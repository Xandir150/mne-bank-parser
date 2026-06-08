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

    public LoaderWorker(LoaderConfig config, Com1CConnector com,
        ILogger<LoaderWorker> logger)
    {
        _config = config;
        _com = com;
        _logger = logger;
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
                "Edit accounts.config.json and restart to apply changes.");
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
                    "Edit accounts.config.json and restart the service to apply changes.");
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

    private void ScanAndLoad()
    {
        if (!Directory.Exists(_config.OutputDir)) return;

        // Check for admin-initiated triggers
        CheckRescanTrigger();
        CheckLookupTrigger();
        CheckAuditTrigger();
        CheckFixTrigger();
        CheckFixCurrencyTrigger();

        var rawFiles = Directory.GetFiles(_config.OutputDir, "*.txt",
            SearchOption.AllDirectories);

        // Group files by their target database so consecutive files reuse the same 1C session.
        // Files whose account isn't mapped go last (they'll error with "not found" anyway).
        var sortedFiles = rawFiles
            .Select(f => new { Path = f, Db = TryFindDatabase(f) })
            .OrderBy(x => x.Db ?? "~~unmapped")
            .Select(x => x.Path)
            .ToArray();

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
                    _logger.LogInformation("All {Skipped} docs in {File} already exist, marking as loaded",
                        result.Skipped, fileName);
                    File.WriteAllText(filePath + ".loaded",
                        $"AllSkipped: {DateTime.Now:yyyy-MM-dd HH:mm:ss}, " +
                        $"Skipped: {result.Skipped}");
                }
                else if (result.Error != null)
                {
                    // Real error — rename to .error so we don't retry forever
                    _logger.LogWarning("Error loading {File}: {Error}", fileName, result.Error);
                    try { File.Move(filePath, filePath + ".error"); } catch { }
                }
                else
                {
                    // Created=0, Skipped=0, no error — empty result, skip
                    _logger.LogWarning("No result for {File}, skipping", fileName);
                    File.WriteAllText(filePath + ".loaded",
                        $"Empty: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing {File}", filePath);
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
    }

    /// <summary>Quickly extract the account number from a 1C bank exchange file and
    /// resolve its target database via the loaded mapping. Returns null on any failure
    /// (file unreadable, account unmapped). Used only for sort ordering — failures are
    /// handled later by the actual load path.</summary>
    private string? TryFindDatabase(string filePath)
    {
        try
        {
            using var sr = new StreamReader(filePath, System.Text.Encoding.GetEncoding(1251));
            string? line;
            int n = 0;
            while ((line = sr.ReadLine()) != null && n++ < 30)
            {
                if (line.StartsWith("РасчСчет=", StringComparison.Ordinal))
                {
                    var acct = line.Substring("РасчСчет=".Length).Trim();
                    return _com.FindAccount(acct)?.Database;
                }
            }
        }
        catch { }
        return null;
    }
}
