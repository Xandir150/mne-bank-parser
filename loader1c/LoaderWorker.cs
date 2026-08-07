using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Hosting;

// ──────────────────────────────────────────────────────────────
// Background worker service
// ──────────────────────────────────────────────────────────────
public class LoaderWorker : BackgroundService
{
    private readonly LoaderConfig _config;
    private readonly Com1CConnector _com;
    private readonly ILogger<LoaderWorker> _logger;

    // Per-file transient-retry state (in-memory: a service restart resets it,
    // which only means a file starts again from attempt 0 — never loses data,
    // retries are idempotent thanks to CountExistingDocs dedupe).
    private readonly Dictionary<string, (int Attempts, DateTime NextTry)> _retryState = new();

    // Discovery caches: accounts that a full sweep failed to find (with TTL),
    // and accounts already suggested to the human (waiting for a config edit).
    // Both reset when accounts.config.json changes or discover.trigger is dropped.
    private readonly Dictionary<string, DateTime> _discoveryMisses = new();
    private readonly HashSet<string> _discoverySuggested = new();
    private long _discoverySeenGen = -1;

    public LoaderWorker(LoaderConfig config, Com1CConnector com,
        ILogger<LoaderWorker> logger)
    {
        _config = config;
        _com = com;
        _logger = logger;
    }

    /// <summary>Move a file to .error and write the cause sidecar (.error.info)
    /// used later by auto-requeue and by humans diagnosing the failure.</summary>
    private void MarkPermanentError(string filePath, string error)
    {
        _logger.LogWarning("Error loading {File}: {Error}", Path.GetFileName(filePath), error);
        try { File.Move(filePath, filePath + ".error"); } catch { }
        try
        {
            File.WriteAllText(filePath + ".error.info",
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}\n{error}\nrequeued: 0\n",
                System.Text.Encoding.UTF8);
        }
        catch { }
        _retryState.Remove(filePath);
    }

    /// <summary>Handle a failed load: transient errors stay .txt and back off
    /// exponentially; permanent errors (or exhausted retries) go to .error.</summary>
    private void HandleLoadError(string filePath, string error)
    {
        bool transient = _config.MaxTransientAttempts > 0 && ErrorClassifier.IsTransient(error);
        if (!transient)
        {
            MarkPermanentError(filePath, error);
            return;
        }

        var (attempts, _) = _retryState.TryGetValue(filePath, out var st) ? st : (0, DateTime.MinValue);
        attempts++;
        if (attempts >= _config.MaxTransientAttempts)
        {
            MarkPermanentError(filePath,
                $"{error}\n(gave up after {attempts} transient attempts)");
            return;
        }

        int cycles = Math.Min(1 << attempts, _config.MaxRetryBackoffCycles);
        var nextTry = DateTime.Now.AddSeconds((double)_config.ScanIntervalSec * cycles);
        _retryState[filePath] = (attempts, nextTry);
        _logger.LogWarning("Transient error on {File} (attempt {N}/{Max}), retry after {Next:HH:mm:ss}: {Error}",
            Path.GetFileName(filePath), attempts, _config.MaxTransientAttempts, nextTry, error);
    }

    /// <summary>Force a thorough GC including LOH compaction. Helps return memory to OS
    /// after batches of file loads create large transient COM RCW chains.</summary>
    private static void DeepGC()
    {
        GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("Loader1C service starting...");

        // Load mapping: human-edited config (preferred) → auto-generated mapping → scan.
        bool hasCached = _com.LoadCachedMapping();
        if (hasCached && _com.IsConfigBased)
        {
            _logger.LogInformation("Mapping loaded from human-edited config — auto-scan disabled. " +
                (_config.ConfigReloadEnabled
                    ? $"Edits to accounts.config.json are picked up automatically within ~{_config.ScanIntervalSec}s."
                    : "Edit accounts.config.json and restart to apply changes."));
        }
        else if (hasCached)
        {
            var ageHours = _com.MappingAgeHours;
            if (_config.RescanAfterHours > 0 && ageHours >= _config.RescanAfterHours)
            {
                _logger.LogInformation("Cached mapping is {Age:F1}h old (>= {Threshold}h), background rescan will run",
                    ageHours, _config.RescanAfterHours);
                _ = Task.Run(() => { try { _com.ScanDatabases(); } catch (Exception ex) { _logger.LogError(ex, "Background rescan failed"); } }, stoppingToken);
            }
            else
            {
                _logger.LogInformation("Cached mapping is fresh ({Age:F1}h old) — skipping rescan",
                    ageHours);
            }
        }
        else
        {
            _logger.LogInformation("No cached mapping, scanning databases...");
            await Task.Run(() => _com.ScanDatabases(), stoppingToken);
        }

        _logger.LogInformation("Watching {Dir} (interval: {Sec}s)",
            _config.OutputDir, _config.ScanIntervalSec);

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                ScanAndLoad();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in scan cycle");
            }

            await Task.Delay(_config.ScanIntervalSec * 1000, stoppingToken);
        }
    }

    /// <summary>If the user (admin) created a sentinel file, kick off a background DB rescan
    /// so newly-added 1C accounts get picked up without restarting the service. The sentinel
    /// is deleted once the scan is launched. No-op when the mapping is config-based.</summary>
    private bool CheckRescanTrigger()
    {
        try
        {
            var dataDir = Path.GetDirectoryName(_config.OutputDir);
            if (string.IsNullOrEmpty(dataDir)) return false;
            var trigger = Path.Combine(dataDir, "rescan.trigger");
            if (!File.Exists(trigger)) return false;
            try { File.Delete(trigger); } catch { }
            if (_com.IsConfigBased)
            {
                _logger.LogWarning("rescan.trigger ignored — mapping is config-based. " +
                    "Edit accounts.config.json — changes are picked up automatically.");
                return false;
            }
            _logger.LogInformation("rescan.trigger detected — starting background DB rescan");
            try { _com.EndSession(); } catch { }
            Task.Run(() => {
                try { _com.ScanDatabases(); }
                catch (Exception ex) { _logger.LogError(ex, "Triggered rescan failed"); }
            });
            return true;
        }
        catch { return false; }
    }

    /// <summary>If the user dropped a <c>lookup.trigger</c> file with content
    /// "<c>&lt;db&gt; &lt;name_substring&gt;</c>", run a targeted org lookup in 1C and
    /// write the result to <c>lookup.result.txt</c>. Used to bypass the COM permission
    /// barrier — only the service account (USR1CV8) can talk to 1C, so external admins
    /// trigger via filesystem.</summary>
    private bool CheckLookupTrigger()
    {
        try
        {
            var dataDir = Path.GetDirectoryName(_config.OutputDir);
            if (string.IsNullOrEmpty(dataDir)) return false;
            var trigger = Path.Combine(dataDir, "lookup.trigger");
            if (!File.Exists(trigger)) return false;

            string content;
            try { content = File.ReadAllText(trigger, System.Text.Encoding.UTF8).Trim(); }
            catch (Exception ex)
            {
                _logger.LogWarning("lookup.trigger: can't read — {Err}", ex.Message);
                return false;
            }
            try { File.Delete(trigger); } catch { }

            var parts = content.Split(new[] { ' ', '\t', '\r', '\n' },
                2, StringSplitOptions.RemoveEmptyEntries);
            var resultFile = Path.Combine(dataDir, "lookup.result.txt");
            if (parts.Length < 2)
            {
                var msg = $"lookup.trigger: expected '<db> <name_substring>', got '{content}'";
                _logger.LogWarning(msg);
                try { File.WriteAllText(resultFile, msg, System.Text.Encoding.UTF8); } catch { }
                return false;
            }

            string db = parts[0];
            string nameFilter = parts[1];
            _logger.LogInformation("Lookup triggered: db={Db} name='{Name}'", db, nameFilter);

            // End any active session first so the lookup gets a fresh connection
            try { _com.EndSession(); } catch { }

            Task.Run(() => {
                try
                {
                    var report = _com.LookupOrg(db, nameFilter);
                    File.WriteAllText(resultFile, report, System.Text.Encoding.UTF8);
                    _logger.LogInformation("Lookup result written to {File}", resultFile);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Lookup failed");
                    try { File.WriteAllText(resultFile, $"ERROR: {ex.Message}\n{ex}",
                        System.Text.Encoding.UTF8); } catch { }
                }
            });
            return true;
        }
        catch { return false; }
    }

    /// <summary>diagnoseorg.trigger, content "&lt;db&gt; &lt;bankAccount&gt;" — runs
    /// DiagnoseOrgWrite to check whether a bank account's Владелец is correctly typed as
    /// Организация (writes+cleans up a throwaway test document). Writes diagnoseorg.result.txt.</summary>
    private bool CheckDiagnoseOrgTrigger()
    {
        try
        {
            var dataDir = Path.GetDirectoryName(_config.OutputDir);
            if (string.IsNullOrEmpty(dataDir)) return false;
            var trigger = Path.Combine(dataDir, "diagnoseorg.trigger");
            if (!File.Exists(trigger)) return false;

            string content;
            try { content = File.ReadAllText(trigger, System.Text.Encoding.UTF8).Trim(); }
            catch { return false; }
            try { File.Delete(trigger); } catch { }

            var parts = content.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var resultFile = Path.Combine(dataDir, "diagnoseorg.result.txt");
            if (parts.Length != 2)
            {
                try { File.WriteAllText(resultFile, $"expected '<db> <account>', got '{content}'", System.Text.Encoding.UTF8); } catch { }
                return false;
            }

            try { _com.EndSession(); } catch { }
            Task.Run(() => {
                try
                {
                    var report = _com.DiagnoseOrgWrite(parts[0], parts[1]);
                    File.WriteAllText(resultFile, report, System.Text.Encoding.UTF8);
                    _logger.LogInformation("diagnoseorg.result written to {File}", resultFile);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "diagnoseorg failed");
                    try { File.WriteAllText(resultFile, $"ERROR: {ex.Message}\n{ex}", System.Text.Encoding.UTF8); } catch { }
                }
            });
            return true;
        }
        catch { return false; }
    }

    /// <summary>inspect.trigger: content "&lt;db&gt; &lt;bankAccount&gt; &lt;dd.MM.yyyy&gt;" —
    /// dumps the actual Организация/Контрагент/etc field values of documents matching that
    /// bank account and date, so you can see why a loaded document "doesn't show up right"
    /// in the 1C UI. Writes inspect.result.txt.</summary>
    private bool CheckInspectTrigger()
    {
        try
        {
            var dataDir = Path.GetDirectoryName(_config.OutputDir);
            if (string.IsNullOrEmpty(dataDir)) return false;
            var trigger = Path.Combine(dataDir, "inspect.trigger");
            if (!File.Exists(trigger)) return false;

            string content;
            try { content = File.ReadAllText(trigger, System.Text.Encoding.UTF8).Trim(); }
            catch (Exception ex)
            {
                _logger.LogWarning("inspect.trigger: can't read — {Err}", ex.Message);
                return false;
            }
            try { File.Delete(trigger); } catch { }

            var parts = content.Split(new[] { ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries);
            var resultFile = Path.Combine(dataDir, "inspect.result.txt");
            if (parts.Length != 3)
            {
                var msg = $"inspect.trigger: expected '<db> <bankAccount> <dd.MM.yyyy>', got '{content}'";
                _logger.LogWarning(msg);
                try { File.WriteAllText(resultFile, msg, System.Text.Encoding.UTF8); } catch { }
                return false;
            }

            string db = parts[0], acct = parts[1], dateStr = parts[2];
            _logger.LogInformation("Inspect triggered: db={Db} account={Acct} date={Date}", db, acct, dateStr);

            try { _com.EndSession(); } catch { }
            Task.Run(() => {
                try
                {
                    var report = _com.InspectRecentDocs(db, acct, dateStr);
                    File.WriteAllText(resultFile, report, System.Text.Encoding.UTF8);
                    _logger.LogInformation("Inspect result written to {File}", resultFile);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Inspect failed");
                    try { File.WriteAllText(resultFile, $"ERROR: {ex.Message}\n{ex}",
                        System.Text.Encoding.UTF8); } catch { }
                }
            });
            return true;
        }
        catch { return false; }
    }

    /// <summary>If the user dropped an <c>audit.trigger</c> file, run a read-only audit of all
    /// databases (or only those listed in the trigger, space/comma-separated) and write
    /// <c>audit.json</c>. Runs under the service account so 1C COM is available.</summary>
    private bool CheckAuditTrigger()
    {
        try
        {
            var dataDir = Path.GetDirectoryName(_config.OutputDir);
            if (string.IsNullOrEmpty(dataDir)) return false;
            var trigger = Path.Combine(dataDir, "audit.trigger");
            if (!File.Exists(trigger)) return false;

            string content = "";
            try { content = File.ReadAllText(trigger, System.Text.Encoding.UTF8).Trim(); } catch { }
            try { File.Delete(trigger); } catch { }

            List<string>? onlyDbs = null;
            if (!string.IsNullOrWhiteSpace(content))
                onlyDbs = content.Split(new[] { ' ', ',', '\t', '\r', '\n' },
                    StringSplitOptions.RemoveEmptyEntries).ToList();

            var outPath = Path.Combine(dataDir, "audit.json");
            _logger.LogInformation("audit.trigger detected — running read-only audit ({Scope})",
                onlyDbs == null ? "all DBs" : string.Join(",", onlyDbs));

            try { _com.EndSession(); } catch { }
            Task.Run(() => {
                try { _com.AuditAll(outPath, onlyDbs); }
                catch (Exception ex) { _logger.LogError(ex, "Audit failed"); }
            });
            return true;
        }
        catch { return false; }
    }

    /// <summary>If the user dropped a <c>fix.trigger</c> file, apply account-number fixes.
    /// Each non-empty, non-'#' line is "<c>db|oldAccount|newAccount</c>". Writes a detailed
    /// before/after report to <c>fix.result.txt</c>. Each fix is collision-guarded and verified.</summary>
    private bool CheckFixTrigger()
    {
        try
        {
            var dataDir = Path.GetDirectoryName(_config.OutputDir);
            if (string.IsNullOrEmpty(dataDir)) return false;
            var trigger = Path.Combine(dataDir, "fix.trigger");
            if (!File.Exists(trigger)) return false;

            string content = "";
            try { content = File.ReadAllText(trigger, System.Text.Encoding.UTF8); } catch { }
            try { File.Delete(trigger); } catch { }

            var lines = content.Split('\n')
                .Select(l => l.Trim())
                .Where(l => l.Length > 0 && !l.StartsWith("#"))
                .ToList();
            var resultFile = Path.Combine(dataDir, "fix.result.txt");
            _logger.LogInformation("fix.trigger detected — {Count} fix line(s)", lines.Count);

            try { _com.EndSession(); } catch { }
            Task.Run(() => {
                var sb = new System.Text.StringBuilder();
                foreach (var line in lines)
                {
                    var p = line.Split('|');
                    if (p.Length != 3)
                    {
                        sb.AppendLine($"SKIP malformed line: '{line}'  (expected db|old|new)");
                        continue;
                    }
                    try { sb.AppendLine(_com.FixAccountNumber(p[0].Trim(), p[1].Trim(), p[2].Trim())); }
                    catch (Exception ex) { sb.AppendLine($"ERROR on '{line}': {ex.Message}"); }
                }
                try { File.WriteAllText(resultFile, sb.ToString(), System.Text.Encoding.UTF8); } catch { }
                _logger.LogInformation("fix.result written to {File}", resultFile);
            });
            return true;
        }
        catch { return false; }
    }

    /// <summary>If the user dropped a <c>fixcur.trigger</c> file, apply currency rename/recode fixes.
    /// Each non-empty, non-'#' line is "<c>db|currentName|newName|newCode</c>". Writes a detailed
    /// before/after report to <c>fixcur.result.txt</c>. Each fix is collision-guarded and verified.</summary>
    private bool CheckFixCurrencyTrigger()
    {
        try
        {
            var dataDir = Path.GetDirectoryName(_config.OutputDir);
            if (string.IsNullOrEmpty(dataDir)) return false;
            var trigger = Path.Combine(dataDir, "fixcur.trigger");
            if (!File.Exists(trigger)) return false;

            string content = "";
            try { content = File.ReadAllText(trigger, System.Text.Encoding.UTF8); } catch { }
            try { File.Delete(trigger); } catch { }

            var lines = content.Split('\n')
                .Select(l => l.Trim())
                .Where(l => l.Length > 0 && !l.StartsWith("#"))
                .ToList();
            var resultFile = Path.Combine(dataDir, "fixcur.result.txt");
            _logger.LogInformation("fixcur.trigger detected — {Count} fix line(s)", lines.Count);

            try { _com.EndSession(); } catch { }
            Task.Run(() => {
                var sb = new System.Text.StringBuilder();
                foreach (var line in lines)
                {
                    var p = line.Split('|');
                    if (p.Length != 4)
                    {
                        sb.AppendLine($"SKIP malformed line: '{line}'  (expected db|curName|newName|newCode)");
                        continue;
                    }
                    try { sb.AppendLine(_com.FixCurrency(p[0].Trim(), p[1].Trim(), p[2].Trim(), p[3].Trim())); }
                    catch (Exception ex) { sb.AppendLine($"ERROR on '{line}': {ex.Message}"); }
                }
                try { File.WriteAllText(resultFile, sb.ToString(), System.Text.Encoding.UTF8); } catch { }
                _logger.LogInformation("fixcur.result written to {File}", resultFile);
            });
            return true;
        }
        catch { return false; }
    }

    /// <summary>markdelete.trigger: cleanup tool for documents created from a malformed source
    /// file. Line format: "<c>db|bankAccount|minAmount|dateFrom|dateTo|dryrun</c>" (dryrun is
    /// "true" or "false" — ALWAYS run true first and inspect the result before a real run).
    /// Writes markdelete.result.txt.</summary>
    private bool CheckMarkDeleteTrigger()
    {
        try
        {
            var dataDir = Path.GetDirectoryName(_config.OutputDir);
            if (string.IsNullOrEmpty(dataDir)) return false;
            var trigger = Path.Combine(dataDir, "markdelete.trigger");
            if (!File.Exists(trigger)) return false;

            string content = "";
            try { content = File.ReadAllText(trigger, System.Text.Encoding.UTF8); } catch { }
            try { File.Delete(trigger); } catch { }

            var lines = content.Split('\n')
                .Select(l => l.Trim())
                .Where(l => l.Length > 0 && !l.StartsWith("#"))
                .ToList();
            var resultFile = Path.Combine(dataDir, "markdelete.result.txt");
            _logger.LogInformation("markdelete.trigger detected — {Count} line(s)", lines.Count);

            try { _com.EndSession(); } catch { }
            Task.Run(() => {
                var sb = new System.Text.StringBuilder();
                foreach (var line in lines)
                {
                    var p = line.Split('|');
                    if (p.Length != 6 || !decimal.TryParse(p[2].Trim(), System.Globalization.NumberStyles.Any,
                            System.Globalization.CultureInfo.InvariantCulture, out var minAmount) ||
                        !bool.TryParse(p[5].Trim(), out var dryRun))
                    {
                        sb.AppendLine($"SKIP malformed line: '{line}'  (expected db|account|minAmount|dateFrom|dateTo|dryrun)");
                        continue;
                    }
                    try { sb.AppendLine(_com.MarkDeleteGarbageDocs(p[0].Trim(), p[1].Trim(), minAmount, p[3].Trim(), p[4].Trim(), dryRun)); }
                    catch (Exception ex) { sb.AppendLine($"ERROR on '{line}': {ex.Message}"); }
                }
                try { File.WriteAllText(resultFile, sb.ToString(), System.Text.Encoding.UTF8); } catch { }
                _logger.LogInformation("markdelete.result written to {File}", resultFile);
            });
            return true;
        }
        catch { return false; }
    }

    private void ScanAndLoad()
    {
        if (!Directory.Exists(_config.OutputDir)) return;

        // Pick up accounts.config.json edits without restart; on change, put
        // .error files whose account became mapped back into the queue.
        if (_config.ConfigReloadEnabled && _com.ReloadAccountConfigIfChanged())
            RequeueErrorFiles();

        // Check for admin-initiated triggers
        CheckRescanTrigger();
        CheckLookupTrigger();
        CheckInspectTrigger();
        CheckDiagnoseOrgTrigger();
        CheckAuditTrigger();
        CheckFixTrigger();
        CheckFixCurrencyTrigger();
        CheckDiscoverTrigger();
        CheckMarkDeleteTrigger();

        var rawFiles = Directory.GetFiles(_config.OutputDir, "*.txt",
            SearchOption.AllDirectories);

        // Group files by their target database so consecutive files reuse the same 1C session.
        // Files whose account isn't mapped go last (they'll error with "not found" anyway).
        var sortedFiles = rawFiles
            .Select(f => new { Path = f, Db = TryFindDatabase(f) })
            .OrderBy(x => x.Db ?? "~~unmapped")
            .Select(x => x.Path)
            .ToArray();

        // Discovery caches follow the mapping generation: a config edit means the
        // human may have added accounts in 1C too — start fresh.
        if (_discoverySeenGen != _com.AccountConfigGeneration)
        {
            _discoverySeenGen = _com.AccountConfigGeneration;
            _discoveryMisses.Clear();
            _discoverySuggested.Clear();
        }
        // TTL expiry for negative cache
        var missTtl = TimeSpan.FromHours(Math.Max(1, _config.DiscoveryNegativeCacheHours));
        foreach (var k in _discoveryMisses.Where(kv => DateTime.Now - kv.Value > missTtl)
                     .Select(kv => kv.Key).ToList())
            _discoveryMisses.Remove(k);

        // Unknown accounts collected this cycle for a single discovery sweep
        var pendingUnknown = new List<(string Acct, string FilePath, string Error)>();

        bool anyProcessed = false;
        foreach (var filePath in sortedFiles)
        {
            var fileName = Path.GetFileName(filePath);

            // Skip files that look like garbage (unknown_*)
            if (fileName.StartsWith("unknown_")) continue;

            // Only skip if .loaded marker exists — delete marker to re-process
            if (File.Exists(filePath + ".loaded")) continue;

            // Re-check file exists (may have been moved between GetFiles and now)
            if (!File.Exists(filePath)) continue;

            // In transient-retry backoff? Skip silently until its next-try time.
            if (_retryState.TryGetValue(filePath, out var retry) && retry.NextTry > DateTime.Now)
                continue;

            _logger.LogInformation("New file: {File}", filePath);

            try
            {
                var parsed = BankFileParser.Parse(filePath);
                if (parsed.Documents.Count == 0)
                {
                    _logger.LogWarning("No documents in {File}, skipping", fileName);
                    continue;
                }

                var result = _com.LoadFile(parsed);

                if (result.Created > 0)
                {
                    _retryState.Remove(filePath);
                    _logger.LogInformation("Loaded {File}: {Created} created, {Skipped} skipped",
                        fileName, result.Created, result.Skipped);

                    // Write .loaded marker (delete this file to re-process)
                    File.WriteAllText(filePath + ".loaded",
                        $"Loaded: {DateTime.Now:yyyy-MM-dd HH:mm:ss}, " +
                        $"Created: {result.Created}, Skipped: {result.Skipped}");

                    // Move to loaded dir (preserve subfolder structure)
                    var relDir = Path.GetRelativePath(_config.OutputDir,
                        Path.GetDirectoryName(filePath)!);
                    var destDir = Path.Combine(_config.LoadedDir, relDir);
                    Directory.CreateDirectory(destDir);
                    var destPath = Path.Combine(destDir, fileName);
                    if (File.Exists(destPath))
                        destPath = Path.Combine(destDir,
                            Path.GetFileNameWithoutExtension(fileName) +
                            "_" + DateTime.Now.ToString("yyyyMMddHHmmss") +
                            Path.GetExtension(fileName));
                    File.Move(filePath, destPath);
                }
                else if (result.Skipped > 0 && result.Error == null)
                {
                    // All documents already exist in 1C — mark as loaded, don't retry
                    _retryState.Remove(filePath);
                    _logger.LogInformation("All {Skipped} docs in {File} already exist, marking as loaded",
                        result.Skipped, fileName);
                    File.WriteAllText(filePath + ".loaded",
                        $"AllSkipped: {DateTime.Now:yyyy-MM-dd HH:mm:ss}, " +
                        $"Skipped: {result.Skipped}");
                }
                else if (result.Error != null)
                {
                    bool isUnknownAccount = result.Error.Contains("not found in any database");
                    if (isUnknownAccount && _config.AutoDiscovery != "off")
                    {
                        var acct = Com1CConnector.NormalizeAccount(parsed.Account);
                        if (_discoverySuggested.Contains(acct))
                        {
                            MarkPermanentError(filePath, result.Error +
                                "\n(предложение уже записано в discovered.accounts.txt — добавьте счёт в accounts.config.json)");
                        }
                        else if (_discoveryMisses.ContainsKey(acct))
                        {
                            MarkPermanentError(filePath, result.Error +
                                "\n(discovery sweep по всем базам тоже не нашёл этот счёт)");
                        }
                        else
                        {
                            // Leave the file as .txt; sweep runs after the loop.
                            pendingUnknown.Add((acct, filePath, result.Error));
                            _logger.LogInformation("Unknown account {Acct} in {File} — queued for discovery sweep",
                                acct, fileName);
                        }
                    }
                    else
                    {
                        // Transient errors stay .txt with backoff; permanent go .error
                        HandleLoadError(filePath, result.Error);
                    }
                }
                else
                {
                    // Created=0, Skipped=0, no error — empty result, skip
                    _retryState.Remove(filePath);
                    _logger.LogWarning("No result for {File}, skipping", fileName);
                    File.WriteAllText(filePath + ".loaded",
                        $"Empty: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                }
            }
            catch (Exception ex)
            {
                // Parse/IO exceptions are treated as transient too: a half-copied
                // file from the parser side often becomes readable a cycle later.
                _logger.LogError(ex, "Error processing {File}", filePath);
                HandleLoadError(filePath, "COM error: " + ex.Message);
            }

            anyProcessed = true;

            // Memory threshold guard: if working set grew past the limit, exit with
            // a non-zero code — Windows Service Manager (with `sc failure ... actions=restart/...`)
            // will bring the service back fresh. Prevents unbounded growth from native 1C COM
            // buffers that the .NET GC can't reclaim.
            if (_config.MaxWorkingSetMB > 0)
            {
                long wsMb = Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024);
                if (wsMb > _config.MaxWorkingSetMB)
                {
                    _logger.LogWarning("WorkingSet {WsMb} MB exceeded threshold {Limit} MB — exiting for service restart",
                        wsMb, _config.MaxWorkingSetMB);
                    // Environment.Exit runs finalizers (releases COM RCWs) then exits with code 2.
                    // SCM sees a failure and triggers recovery → fresh process.
                    Environment.Exit(2);
                }
            }
        }

        // End of cycle: release the 1C session so memory can settle, then deep-GC.
        if (anyProcessed)
        {
            try { _com.EndSession(); } catch { }
            DeepGC();
        }

        // One discovery sweep per cycle for accounts unknown to the mapping.
        if (pendingUnknown.Count > 0 && _config.MaxDiscoverySweepsPerCycle > 0)
            RunDiscoverySweep(pendingUnknown);
    }

    /// <summary>discover.trigger: clear the discovery caches so unknown accounts are
    /// swept again immediately (e.g. after the accountant added the account in 1C
    /// without touching accounts.config.json).</summary>
    private bool CheckDiscoverTrigger()
    {
        try
        {
            var dataDir = Path.GetDirectoryName(_config.OutputDir);
            if (string.IsNullOrEmpty(dataDir)) return false;
            var trigger = Path.Combine(dataDir, "discover.trigger");
            if (!File.Exists(trigger)) return false;
            try { File.Delete(trigger); } catch { }
            _discoveryMisses.Clear();
            _discoverySuggested.Clear();
            _logger.LogInformation("discover.trigger detected — discovery caches cleared, unknown accounts will be swept again");
            RequeueErrorFiles(requireMapped: false);
            return true;
        }
        catch { return false; }
    }

    /// <summary>Sweep all DBs for the cycle's unknown accounts; per account either
    /// write a config suggestion (suggest mode), auto-append (append mode), or
    /// negative-cache the miss. Files stay .txt and resolve on a following cycle.</summary>
    private void RunDiscoverySweep(List<(string Acct, string FilePath, string Error)> pending)
    {
        try
        {
            var accounts = pending.Select(p => p.Acct).Distinct().ToList();
            var hits = _com.TryDiscoverAccounts(accounts);
            var dataDir = Path.GetDirectoryName(_config.OutputDir) ?? "";
            var suggestFile = Path.Combine(dataDir, "discovered.accounts.txt");

            foreach (var acct in accounts)
            {
                var acctHits = hits.Where(h => h.Account == acct).ToList();
                if (acctHits.Count == 0)
                {
                    _discoveryMisses[acct] = DateTime.Now;
                    _logger.LogWarning("DISCOVERY: account {Acct} not found in any of {Count} DBs — will .error next cycle",
                        acct, _config.Databases.Count);
                    continue;
                }

                bool ambiguous = acctHits.Select(h => h.Database).Distinct().Count() > 1;
                var lines = new List<string>
                {
                    $"// ─── {DateTime.Now:yyyy-MM-dd HH:mm:ss} " +
                    (ambiguous ? "— НЕОДНОЗНАЧНО: счёт найден в нескольких базах, выберите одну ───"
                               : "───"),
                };
                foreach (var h in acctHits)
                    lines.Add($"  \"{h.Database}\": {{ \"account\": \"{h.Account}\", \"org\": \"{h.OrgName.Replace("\"", "\\\"")}\", \"inn\": \"{h.Inn}\" }}");
                try { File.AppendAllLines(suggestFile, lines, System.Text.Encoding.UTF8); } catch { }

                if (!ambiguous && _config.AutoDiscovery == "append")
                {
                    var h = acctHits[0];
                    if (_com.AppendAccountConfigEntry(h.Database, h.Account, h.OrgName, h.Inn))
                        continue; // mapped now — the .txt file loads next cycle
                }

                _discoverySuggested.Add(acct);
                _logger.LogWarning("DISCOVERY: account {Acct} found in {Dbs} — снипет записан в discovered.accounts.txt, добавьте в accounts.config.json",
                    acct, string.Join(",", acctHits.Select(h => h.Database).Distinct()));
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Discovery sweep failed");
        }
    }

    /// <summary>Quickly extract the account number from a 1C bank exchange file and
    /// resolve its target database via the loaded mapping. Returns null on any failure
    /// (file unreadable, account unmapped). Used only for sort ordering — failures are
    /// handled later by the actual load path.</summary>
    private string? TryFindDatabase(string filePath)
    {
        var acct = TryReadAccount(filePath);
        return acct == null ? null : _com.FindAccount(acct)?.Database;
    }

    /// <summary>Extract РасчСчет from the first lines of a 1C bank exchange file.
    /// Null on any failure. Works for .txt and .error files alike.</summary>
    private static string? TryReadAccount(string filePath)
    {
        try
        {
            using var sr = new StreamReader(filePath, System.Text.Encoding.GetEncoding(1251));
            string? line;
            int n = 0;
            while ((line = sr.ReadLine()) != null && n++ < 30)
            {
                if (line.StartsWith("РасчСчет=", StringComparison.Ordinal))
                    return line.Substring("РасчСчет=".Length).Trim();
            }
        }
        catch { }
        return null;
    }

    /// <summary>After the account mapping changed, put matching .error files back
    /// into the queue automatically. Only files whose failure cause was
    /// "account not found" AND whose account is now mapped are re-queued;
    /// each file is re-queued at most MaxAutoRequeues times (tracked in the
    /// .error.info sidecar) to prevent .txt ↔ .error cycling.</summary>
    private void RequeueErrorFiles(bool requireMapped = true)
    {
        if (!_config.AutoRequeueErrors) return;
        try
        {
            foreach (var errFile in Directory.GetFiles(_config.OutputDir, "*.txt.error",
                SearchOption.AllDirectories))
            {
                try
                {
                    var infoFile = errFile + ".info"; // <file>.txt.error.info
                    string info = "";
                    bool hasInfo = File.Exists(infoFile);
                    if (hasInfo)
                        info = File.ReadAllText(infoFile, System.Text.Encoding.UTF8);

                    // Cause check: sidecar says unknown-account, or (no sidecar,
                    // legacy .error file) — fall through to the account check directly.
                    if (hasInfo && !info.Contains("not found in any database"))
                        continue;

                    // Requeue budget
                    int requeued = 0;
                    var m = System.Text.RegularExpressions.Regex.Match(info, @"requeued:\s*(\d+)");
                    if (m.Success) requeued = int.Parse(m.Groups[1].Value);
                    if (requeued >= _config.MaxAutoRequeues) continue;

                    // Is the account mapped now? (discover.trigger re-queues even
                    // unmapped ones so they go through a fresh discovery sweep)
                    var acct = TryReadAccount(errFile);
                    if (acct == null) continue;
                    var mapping = _com.FindAccount(acct);
                    if (mapping == null && requireMapped) continue;

                    var txtPath = errFile.Substring(0, errFile.Length - ".error".Length);
                    if (File.Exists(txtPath)) continue; // shouldn't happen; be safe
                    File.Move(errFile, txtPath);
                    try
                    {
                        var updated = m.Success
                            ? System.Text.RegularExpressions.Regex.Replace(info,
                                @"requeued:\s*\d+", $"requeued: {requeued + 1}")
                            : info + $"\nrequeued: {requeued + 1}\n";
                        File.WriteAllText(infoFile, updated, System.Text.Encoding.UTF8);
                    }
                    catch { }
                    _logger.LogInformation("Re-queued {File} — account {Acct} {Why}",
                        Path.GetFileName(txtPath), acct,
                        mapping != null ? $"is now mapped to {mapping.Database}" : "will go through a fresh discovery sweep");
                }
                catch (Exception ex)
                {
                    _logger.LogWarning("Requeue check failed for {File}: {Err}", errFile, ex.Message);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning("RequeueErrorFiles failed: {Err}", ex.Message);
        }
    }
}
