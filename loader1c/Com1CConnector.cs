using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Logging;

// ──────────────────────────────────────────────────────────────
// 1C COM connector — all COM calls run on a dedicated STA thread
// ──────────────────────────────────────────────────────────────
public partial class Com1CConnector : IDisposable
{
    private readonly LoaderConfig _config;
    private readonly ILogger _logger;

    // Account number → mapping
    private Dictionary<string, AccountMapping> _accountMap = new();
    private readonly string _mappingFile;        // auto-generated, machine-only
    private readonly string _reportFile;
    private readonly string _accountConfigFile;  // human-edited, authoritative when present
    // True when the active mapping was loaded from the human-edited config (skip auto-scan).
    public bool IsConfigBased { get; private set; }

    // Shared payroll document reference for current file being loaded
    private dynamic? _currentPayrollRef;
    private dynamic? _cachedDeptRef;
    private string _cachedDeptName = "";

    // ── Per-session COM caches (persist across files within the same database) ──
    // The connection + caches live for as long as consecutive files target the same DB.
    // Goal: reuse 1C session and catalog lookups instead of re-querying per file.
    private readonly Dictionary<string, dynamic?> _cacheChartAccts = new();
    private readonly Dictionary<string, dynamic?> _cacheBankAcctsByNumber = new();
    private readonly Dictionary<string, dynamic?> _cacheCounterpartyByInn = new();
    private readonly Dictionary<string, dynamic?> _cacheCounterpartyByName = new();
    private readonly Dictionary<string, dynamic?> _cacheBankRefByCode = new();
    private readonly Dictionary<string, dynamic?> _cacheEnumValues = new();
    private readonly Dictionary<string, dynamic?> _cacheCashFlowItems = new();
    // org name → (deptRef, deptName); cached per session (org membership is DB-wide)
    private readonly Dictionary<string, (dynamic? Ref, string Name)> _cacheDeptByOrg = new();
    private dynamic? _cacheEur;
    private dynamic? _cacheVidSchets;
    private bool _cacheVidSchetsAttempted;
    // RCWs scheduled for release at end of session
    private readonly List<object> _comToRelease = new();
    // Lock for _comConnector to avoid double-init from background ScanDatabases vs LoadFile
    private readonly object _connectorLock = new();

    // ── Active session ──
    // Reused across multiple LoadFile calls when they target the same database.
    private dynamic? _sessionConn;
    private string? _sessionDatabase;

    private void TrackCom(object? obj)
    {
        if (obj != null) _comToRelease.Add(obj);
    }

    private static void SafeRelease(object? obj)
    {
        if (obj == null) return;
        try { Marshal.FinalReleaseComObject(obj); } catch { }
    }

    /// <summary>Reset all session-level caches (called on DB switch or end of batch).</summary>
    private void ResetSessionCache()
    {
        foreach (var r in _comToRelease) SafeRelease(r);
        _comToRelease.Clear();
        _cacheChartAccts.Clear();
        _cacheBankAcctsByNumber.Clear();
        _cacheCounterpartyByInn.Clear();
        _cacheCounterpartyByName.Clear();
        _cacheBankRefByCode.Clear();
        _cacheEnumValues.Clear();
        _cacheCashFlowItems.Clear();
        _cacheDeptByOrg.Clear();
        _cacheEur = null;
        _cacheVidSchets = null;
        _cacheVidSchetsAttempted = false;
    }

    /// <summary>End the current session: release connection + clear caches.
    /// Called by the worker between scan cycles or when no more files to process.</summary>
    public void EndSession()
    {
        ResetSessionCache();
        _currentPayrollRef = null;
        _cachedDeptRef = null;
        _cachedDeptName = "";
        if (_sessionConn != null)
        {
            try { CleanupCom(_sessionConn); } catch { }
            _sessionConn = null;
            _sessionDatabase = null;
        }
    }

    /// <summary>Cached lookup for ПодразделениеОрганизации by org name.</summary>
    private (dynamic? Ref, string Name) GetDeptByOrg(dynamic conn, string orgName)
    {
        if (_cacheDeptByOrg.TryGetValue(orgName, out var cached)) return cached;

        dynamic? sel = null;
        dynamic? deptRef = null;
        string deptName = "";
        try
        {
            sel = conn.Справочники.ПодразделенияОрганизаций.Выбрать();
            while (sel.Следующий())
            {
                if ((bool)sel.ПометкаУдаления) continue;
                try
                {
                    string ownerName = (string)(sel.Владелец?.Наименование ?? "");
                    if (ownerName == orgName)
                    {
                        deptRef = sel.Ссылка;
                        TrackCom(deptRef);
                        deptName = (string)(sel.Наименование ?? "");
                        break;
                    }
                }
                catch { }
            }
        }
        catch { }
        finally { SafeRelease(sel); }

        var result = (deptRef, deptName);
        _cacheDeptByOrg[orgName] = result;
        return result;
    }

    /// <summary>Cached lookup for plan account code (e.g. "51", "60.01").</summary>
    internal dynamic? GetChartAccount(dynamic conn, string code)
    {
        if (string.IsNullOrWhiteSpace(code)) return null;
        if (_cacheChartAccts.TryGetValue(code, out var cached)) return cached;
        dynamic? result = null;
        try
        {
            dynamic acct = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду(code);
            if (!acct.Пустая())
            {
                result = acct;
                TrackCom(acct);
            }
        }
        catch { }
        _cacheChartAccts[code] = result;
        return result;
    }

    /// <summary>Cached lookup for EUR currency.</summary>
    internal dynamic? GetEur(dynamic conn)
    {
        if (_cacheEur != null) return _cacheEur;
        try
        {
            dynamic eur = conn.Справочники.Валюты.НайтиПоНаименованию("EUR", true);
            if (!eur.Пустая())
            {
                _cacheEur = eur;
                TrackCom(eur);
            }
        }
        catch { }
        return _cacheEur;
    }

    /// <summary>Cached lookup for enum value: Перечисления.{enumType}.{valueName}.</summary>
    internal dynamic? GetEnumValue(dynamic conn, string enumType, string valueName)
    {
        string key = enumType + "." + valueName;
        if (_cacheEnumValues.TryGetValue(key, out var cached)) return cached;
        dynamic? val = null;
        try
        {
            dynamic enumMgr = conn.Перечисления.GetType().InvokeMember(
                enumType, System.Reflection.BindingFlags.GetProperty,
                null, (object)conn.Перечисления, null);
            val = enumMgr.GetType().InvokeMember(valueName,
                System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
            if (val != null) TrackCom(val);
        }
        catch { }
        _cacheEnumValues[key] = val;
        return val;
    }

    /// <summary>Cached lookup for СтатьяДвиженияДенежныхСредств by name.</summary>
    internal dynamic? GetCashFlowItem(dynamic conn, string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        if (_cacheCashFlowItems.TryGetValue(name, out var cached)) return cached;
        dynamic? result = null;
        try
        {
            dynamic item = conn.Справочники.СтатьиДвиженияДенежныхСредств
                .НайтиПоНаименованию(name, true);
            if (!item.Пустая())
            {
                result = item;
                TrackCom(item);
            }
        }
        catch { }
        _cacheCashFlowItems[name] = result;
        return result;
    }

    /// <summary>Get a sample ВидСчета value to copy to new bank accounts. Caches once per file.</summary>
    internal dynamic? GetSampleVidSchets(dynamic conn)
    {
        if (_cacheVidSchetsAttempted) return _cacheVidSchets;
        _cacheVidSchetsAttempted = true;

        dynamic? sel = null;
        try
        {
            sel = conn.Справочники.БанковскиеСчета.Выбрать();
            while (sel.Следующий())
            {
                dynamic? obj = null;
                try
                {
                    if ((bool)sel.ПометкаУдаления) continue;
                    obj = sel.Ссылка.ПолучитьОбъект();
                    dynamic vid = obj.ВидСчета;
                    string vidStr = vid?.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(vidStr) && vidStr != "System.__ComObject")
                    {
                        _cacheVidSchets = vid;
                        TrackCom(vid);
                        return _cacheVidSchets;
                    }
                }
                catch { }
                finally { SafeRelease(obj); }
            }
        }
        catch { }
        finally { SafeRelease(sel); }
        return null;
    }

    /// <summary>Cached lookup for Банк by 3-digit code (first 3 digits of account).</summary>
    internal dynamic? GetBankByCode(dynamic conn, string bankCode)
    {
        if (string.IsNullOrWhiteSpace(bankCode)) return null;
        if (_cacheBankRefByCode.TryGetValue(bankCode, out var cached)) return cached;

        dynamic? sel = null;
        dynamic? result = null;
        try
        {
            sel = conn.Справочники.Банки.Выбрать();
            while (sel.Следующий())
            {
                try
                {
                    string code = ((string)(sel.Код ?? "")).Trim();
                    if (code == bankCode)
                    {
                        result = sel.Ссылка;
                        TrackCom(result);
                        break;
                    }
                }
                catch { }
            }
        }
        catch { }
        finally { SafeRelease(sel); }
        _cacheBankRefByCode[bankCode] = result;
        return result;
    }

    public Com1CConnector(LoaderConfig config, ILogger logger)
    {
        _config = config;
        _logger = logger;
        // Store mapping in OutputDir (data directory), not in bin/ which gets overwritten on rebuild
        var dataDir = !string.IsNullOrWhiteSpace(config.OutputDir)
            ? Path.GetDirectoryName(config.OutputDir) ?? AppContext.BaseDirectory
            : AppContext.BaseDirectory;
        _mappingFile = Path.Combine(dataDir, "account_mapping.json");
        _reportFile = Path.Combine(dataDir, "account_mapping.txt");
        _accountConfigFile = Path.Combine(dataDir, "accounts.config.json");
    }

    public IReadOnlyDictionary<string, AccountMapping> AccountMap => _accountMap;

    /// <summary>Age of the cached mapping file in hours (double.MaxValue if not present).</summary>
    public double MappingAgeHours
    {
        get
        {
            try
            {
                if (!File.Exists(_mappingFile)) return double.MaxValue;
                return (DateTime.Now - File.GetLastWriteTime(_mappingFile)).TotalHours;
            }
            catch { return double.MaxValue; }
        }
    }

    // Cached COM connector (reused across connections)
    private dynamic? _comConnector;

    /// <summary>Connect to a specific database.</summary>
    private dynamic Connect(string database)
    {
        lock (_connectorLock)
        {
            if (_comConnector == null)
            {
                var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                    ?? throw new Exception("V83.COMConnector not registered");
                _comConnector = Activator.CreateInstance(connectorType)!;
            }
        }
        var connStr = $"Srvr=\"{_config.Server}\";Ref=\"{database}\";" +
                      $"Usr=\"{_config.User}\";Pwd=\"{_config.Password}\"";
        return _comConnector.Connect(connStr);
    }

    /// <summary>Release COM connection and aggressively collect all COM wrappers.</summary>
    public static void CleanupCom(dynamic conn)
    {
        try { while (Marshal.FinalReleaseComObject(conn) > 0) { } } catch { }
        // Multiple GC passes to release nested COM RCW chains
        for (int i = 0; i < 3; i++)
        {
            GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, true, true);
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>Scan all databases and build account→database mapping.
    /// Additive: if a DB scan fails, existing entries for that DB are preserved.
    /// The mapping file is only overwritten when at least one DB was scanned successfully.</summary>
    public void ScanDatabases()
    {
        _logger.LogInformation("Scanning {Count} databases...", _config.Databases.Count);

        // Start from existing mapping as baseline — we'll only replace entries for successfully-scanned DBs.
        var map = new Dictionary<string, AccountMapping>(_accountMap);
        var successfullyScanned = new HashSet<string>();
        int failedCount = 0;

        foreach (var db in _config.Databases)
        {
            dynamic? conn = null;
            dynamic? query = null;
            dynamic? qResult = null;
            dynamic? sel = null;
            var dbAccounts = new Dictionary<string, AccountMapping>();
            bool dbOk = false;
            try
            {
                _logger.LogInformation("Scanning {Database}...", db);
                conn = Connect(db);

                // Read bank accounts owned by Organizations only (not Counterparties)
                query = conn.NewObject("Запрос");
                query.Текст = @"ВЫБРАТЬ
                    БанковскиеСчета.НомерСчета КАК НомерСчета,
                    БанковскиеСчета.Владелец.Наименование КАК ОргНаименование,
                    БанковскиеСчета.Владелец.ИНН КАК ОргИНН
                ИЗ
                    Справочник.БанковскиеСчета КАК БанковскиеСчета
                ГДЕ
                    БанковскиеСчета.ПометкаУдаления = ЛОЖЬ
                    И ТИПЗНАЧЕНИЯ(БанковскиеСчета.Владелец) = ТИП(Справочник.Организации)";

                qResult = query.Выполнить();
                sel = qResult.Выбрать();
                while (sel.Следующий())
                {
                    try
                    {
                        string acctNo = (string)(sel.НомерСчета ?? "");
                        if (string.IsNullOrWhiteSpace(acctNo)) continue;

                        string normalized = NormalizeAccount(acctNo);
                        if (normalized.Length < 10) continue;

                        string orgName = (string)(sel.ОргНаименование ?? "");
                        string inn = (string)(sel.ОргИНН ?? "");

                        if (!dbAccounts.ContainsKey(normalized))
                        {
                            dbAccounts[normalized] = new AccountMapping
                            {
                                Database = db,
                                OrgName = orgName,
                                Inn = inn
                            };
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning("Error reading account: {Error}", ex.Message);
                    }
                }
                dbOk = true;
            }
            catch (Exception ex)
            {
                failedCount++;
                _logger.LogError("Failed to scan {Database}: {Error}", db, ex.Message);
            }
            finally
            {
                SafeRelease(sel);
                SafeRelease(qResult);
                SafeRelease(query);
                if (conn != null) CleanupCom(conn);
            }

            if (dbOk)
            {
                // Remove old entries for this DB (account may have been deleted in 1C)
                var stale = map.Where(kvp => kvp.Value.Database == db).Select(kvp => kvp.Key).ToList();
                foreach (var k in stale) map.Remove(k);
                // Add freshly-scanned entries
                foreach (var kvp in dbAccounts) map[kvp.Key] = kvp.Value;
                successfullyScanned.Add(db);
            }
        }

        _accountMap = map;
        _logger.LogInformation("Scan complete: {Count} accounts mapped, {Ok}/{Total} DBs scanned successfully",
            map.Count, successfullyScanned.Count, _config.Databases.Count);

        // Refuse to persist a mapping that was clobbered by mass failures.
        // Heuristic: require at least 1 successful DB AND non-empty result.
        if (successfullyScanned.Count == 0 || map.Count == 0)
        {
            _logger.LogWarning("Mapping file NOT saved — {Failed} DBs failed, no successful scans. Existing file preserved.",
                failedCount);
            return;
        }

        // Save JSON mapping
        var json = JsonSerializer.Serialize(map, new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
        File.WriteAllText(_mappingFile, json, Encoding.UTF8);
        _logger.LogInformation("Mapping saved to {File}", _mappingFile);

        // Save human-readable TXT report
        SaveReport(map);
        _logger.LogInformation("Report saved to {File}", _reportFile);
    }

    private void SaveReport(Dictionary<string, AccountMapping> map)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Маппинг счетов → базы 1С (обновлено: {DateTime.Now:dd.MM.yyyy HH:mm})");
        sb.AppendLine(new string('=', 120));
        sb.AppendLine($"{"Счёт",-25} {"Организация",-40} {"ИНН",-12} {"База 1С",-20}");
        sb.AppendLine(new string('-', 120));

        foreach (var kvp in map.OrderBy(x => x.Value.Database).ThenBy(x => x.Value.OrgName))
        {
            sb.AppendLine($"{kvp.Key,-25} {kvp.Value.OrgName,-40} {kvp.Value.Inn,-12} {kvp.Value.Database,-20}");
        }

        sb.AppendLine(new string('=', 120));
        sb.AppendLine($"Всего счетов: {map.Count}");
        File.WriteAllText(_reportFile, sb.ToString(), Encoding.UTF8);
    }

    /// <summary>Load account mapping. Prefers the human-edited config file
    /// (accounts.config.json) when present, otherwise falls back to the auto-generated
    /// account_mapping.json from a previous scan.</summary>
    public bool LoadCachedMapping()
    {
        // 1) Human-edited config takes priority — it's the source of truth.
        if (LoadAccountConfig()) return true;

        // 2) Fall back to auto-generated mapping (legacy / pre-config path).
        if (!File.Exists(_mappingFile)) return false;
        try
        {
            var json = File.ReadAllText(_mappingFile, Encoding.UTF8);
            _accountMap = JsonSerializer.Deserialize<Dictionary<string, AccountMapping>>(json)
                          ?? new();
            IsConfigBased = false;
            _logger.LogInformation("Loaded auto-generated mapping: {Count} accounts (from account_mapping.json)",
                _accountMap.Count);
            return _accountMap.Count > 0;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>Entry in the human-edited accounts.config.json.</summary>
    public class AccountConfigEntry
    {
        public string Account { get; set; } = "";
        public string Org { get; set; } = "";
        public string Inn { get; set; } = "";
    }

    /// <summary>Load accounts.config.json (JSONC: comments + trailing commas allowed).
    /// Structure: <c>{ "db_name": [ { "account": "...", "org": "...", "inn": "..." } ] }</c>.
    /// Returns true if loaded successfully and contains at least one account.</summary>
    public bool LoadAccountConfig()
    {
        if (!File.Exists(_accountConfigFile)) return false;
        try
        {
            var raw = File.ReadAllText(_accountConfigFile, Encoding.UTF8);
            var opts = new JsonSerializerOptions
            {
                ReadCommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true,
                PropertyNameCaseInsensitive = true,
            };
            var grouped = JsonSerializer.Deserialize<Dictionary<string, List<AccountConfigEntry>>>(raw, opts);
            if (grouped == null) return false;

            var map = new Dictionary<string, AccountMapping>();
            int dups = 0;
            foreach (var (db, entries) in grouped)
            {
                // Skip keys starting with _ — reserved for metadata/comments
                if (db.StartsWith("_")) continue;
                if (entries == null) continue;
                foreach (var e in entries)
                {
                    if (string.IsNullOrWhiteSpace(e.Account)) continue;
                    var norm = NormalizeAccount(e.Account);
                    if (norm.Length < 10) continue;
                    if (map.ContainsKey(norm))
                    {
                        dups++;
                        _logger.LogWarning("Duplicate account {Acct} in config (db={Old} vs db={New}) — keeping first",
                            norm, map[norm].Database, db);
                        continue;
                    }
                    map[norm] = new AccountMapping
                    {
                        Database = db,
                        OrgName = e.Org ?? "",
                        Inn = e.Inn ?? ""
                    };
                }
            }
            if (map.Count == 0) return false;
            _accountMap = map;
            IsConfigBased = true;
            _logger.LogInformation("Loaded {Count} accounts from accounts.config.json{Dup}",
                map.Count, dups > 0 ? $" ({dups} duplicates skipped)" : "");
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError("Failed to read accounts.config.json: {Err}", ex.Message);
            return false;
        }
    }

    /// <summary>Save current mapping as a human-editable accounts.config.json file,
    /// grouped by database. Use to seed the config from a previous auto-scan.</summary>
    public void SaveAccountConfig()
    {
        var grouped = _accountMap
            .GroupBy(kvp => kvp.Value.Database)
            .OrderBy(g => g.Key)
            .ToDictionary(
                g => g.Key,
                g => g.Select(kvp => new AccountConfigEntry
                {
                    Account = kvp.Key,
                    Org = kvp.Value.OrgName,
                    Inn = kvp.Value.Inn
                })
                .OrderBy(e => e.Org)
                .ThenBy(e => e.Account)
                .ToList()
            );

        var opts = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        };
        var body = JsonSerializer.Serialize(grouped, opts);

        var header = string.Join('\n', new[]
        {
            "// ┌──────────────────────────────────────────────────────────────────────────┐",
            "// │  accounts.config.json  —  ручной маппинг счетов 1С                       │",
            "// │                                                                          │",
            "// │  Формат: { \"<имя_базы_1С>\": [ { account, org, inn }, ... ] }             │",
            "// │  • account — номер счёта, 18 цифр (без ME25 и тире)                      │",
            "// │  • org     — название организации (свободный текст, для удобства поиска) │",
            "// │  • inn     — ИНН, 8 цифр                                                 │",
            "// │                                                                          │",
            "// │  ПОДДЕРЖИВАЮТСЯ комментарии (//) и trailing commas.                      │",
            "// │  ПОСЛЕ правок ОБЯЗАТЕЛЬНО перезапустите сервис:                          │",
            "// │      sc stop Loader1C && sc start Loader1C                               │",
            "// │                                                                          │",
            "// │  При наличии этого файла авто-скан 1С на старте НЕ запускается.          │",
            "// │  Удалите файл и rescan.trigger чтобы вернуться к авто-сканированию.      │",
            "// └──────────────────────────────────────────────────────────────────────────┘",
            "",
            ""
        });

        File.WriteAllText(_accountConfigFile, header + body, new UTF8Encoding(false));
        _logger.LogInformation("Saved {Count} accounts across {Dbs} databases to {File}",
            _accountMap.Count, grouped.Count, _accountConfigFile);
    }

    /// <summary>Path to the human-editable config file (for diagnostics/CLI).</summary>
    public string AccountConfigPath => _accountConfigFile;

    /// <summary>Query one 1C database for organizations whose name contains <paramref name="nameFilter"/>
    /// (case-insensitive). Returns a multi-line human report + a JSON snippet ready to paste into
    /// accounts.config.json. Used by both the CLI command and the worker's trigger handler so the
    /// query always runs under the service account (which has 1C COM permissions).</summary>
    public string LookupOrg(string database, string nameFilter)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Lookup: '{nameFilter}' in {database}");
        sb.AppendLine($"Time:   {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine();

        dynamic? conn = null;
        dynamic? query = null;
        dynamic? result = null;
        dynamic? sel = null;
        try
        {
            conn = Connect(database);

            query = conn.NewObject("Запрос");
            query.Текст = @"ВЫБРАТЬ
                БанковскиеСчета.НомерСчета КАК НомерСчета,
                БанковскиеСчета.ПометкаУдаления КАК Удалён,
                БанковскиеСчета.Владелец.Наименование КАК ОргНаименование,
                БанковскиеСчета.Владелец.ИНН КАК ОргИНН,
                БанковскиеСчета.Владелец.ПометкаУдаления КАК ОргУдалён
            ИЗ
                Справочник.БанковскиеСчета КАК БанковскиеСчета
            ГДЕ
                ТИПЗНАЧЕНИЯ(БанковскиеСчета.Владелец) = ТИП(Справочник.Организации)
                И ВРЕГ(БанковскиеСчета.Владелец.Наименование) ПОДОБНО &ИмяОрг
            УПОРЯДОЧИТЬ ПО
                БанковскиеСчета.Владелец.Наименование,
                БанковскиеСчета.НомерСчета";
            query.УстановитьПараметр("ИмяОрг", "%" + nameFilter.ToUpper() + "%");

            result = query.Выполнить();
            sel = result.Выбрать();

            var found = new Dictionary<string, (string Inn, bool OrgDeleted, List<(string Acct, bool Deleted)> Accts)>();
            while (sel.Следующий())
            {
                try
                {
                    string orgName = ((string)(sel.ОргНаименование ?? "")).Trim();
                    string inn = ((string)(sel.ОргИНН ?? "")).Trim();
                    bool orgDel = false; try { orgDel = (bool)sel.ОргУдалён; } catch { }
                    string acct = ((string)(sel.НомерСчета ?? "")).Trim();
                    bool acctDel = false; try { acctDel = (bool)sel.Удалён; } catch { }
                    if (string.IsNullOrWhiteSpace(orgName)) continue;

                    if (!found.ContainsKey(orgName))
                        found[orgName] = (inn, orgDel, new List<(string, bool)>());
                    found[orgName].Accts.Add((acct, acctDel));
                }
                catch (Exception ex) { sb.AppendLine($"  row error: {ex.Message}"); }
            }

            if (found.Count == 0)
            {
                sb.AppendLine($"No organizations matching '{nameFilter}' found in {database}.");
                return sb.ToString();
            }

            foreach (var kvp in found)
            {
                sb.AppendLine($"=== {kvp.Key}{(kvp.Value.OrgDeleted ? "  [ОРГ. ПОМЕЧЕНА НА УДАЛЕНИЕ!]" : "")}");
                sb.AppendLine($"    ИНН: {(string.IsNullOrEmpty(kvp.Value.Inn) ? "(пусто)" : kvp.Value.Inn)}");
                sb.AppendLine($"    Счетов: {kvp.Value.Accts.Count}");
                foreach (var (acct, del) in kvp.Value.Accts)
                {
                    var norm = NormalizeAccount(acct);
                    string flag = del ? "  [удалён]" : "";
                    string lenWarn = norm.Length != 18 && !del ? $"  [⚠ длина {norm.Length}, ожидалось 18]" : "";
                    sb.AppendLine($"      {acct,-22} → norm: {norm}{flag}{lenWarn}");
                }
                sb.AppendLine();
            }

            // Ready-to-paste JSON snippet (skips deleted accounts and deleted orgs)
            sb.AppendLine("--- Сниппет для accounts.config.json ---");
            sb.AppendLine($"  \"{database}\": [");
            var rows = new List<string>();
            foreach (var kvp in found)
            {
                if (kvp.Value.OrgDeleted) continue;
                foreach (var (acct, del) in kvp.Value.Accts)
                {
                    if (del) continue;
                    var norm = NormalizeAccount(acct);
                    if (string.IsNullOrWhiteSpace(norm)) continue;
                    var esc = kvp.Key.Replace("\\", "\\\\").Replace("\"", "\\\"");
                    rows.Add($"    {{ \"account\": \"{norm}\", \"org\": \"{esc}\", \"inn\": \"{kvp.Value.Inn}\" }}");
                }
            }
            sb.AppendLine(string.Join(",\n", rows));
            sb.AppendLine("  ],");
        }
        catch (Exception ex)
        {
            sb.AppendLine($"ERROR: {ex.Message}");
            sb.AppendLine(ex.ToString());
        }
        finally
        {
            SafeRelease(sel);
            SafeRelease(result);
            SafeRelease(query);
            if (conn != null) try { CleanupCom(conn); } catch { }
        }

        return sb.ToString();
    }

    /// <summary>Find which database an account belongs to (exact match only).</summary>
    public AccountMapping? FindAccount(string account)
    {
        var normalized = NormalizeAccount(account);
        if (_accountMap.TryGetValue(normalized, out var mapping))
            return mapping;

        return null;
    }

    /// <summary>Load a bank exchange file into 1C.</summary>
    public LoadResult LoadFile(BankExchangeFile file)
    {
        var mapping = FindAccount(file.Account);
        if (mapping == null)
        {
            return new LoadResult
            {
                Success = false,
                Error = $"Account {file.Account} not found in any database"
            };
        }

        _logger.LogInformation("Loading {File} into {Database} ({Org})",
            Path.GetFileName(file.FilePath), mapping.Database, mapping.OrgName);

        // Reuse session if same DB; otherwise close old + open new
        if (_sessionDatabase != mapping.Database || _sessionConn == null)
        {
            EndSession();
            try
            {
                _sessionConn = Connect(mapping.Database);
                _sessionDatabase = mapping.Database;
                _logger.LogInformation("Opened new 1C session for {Db}", mapping.Database);
            }
            catch (Exception ex)
            {
                return new LoadResult { Success = false, Error = $"Connect failed: {ex.Message}" };
            }
        }

        // Per-file state reset (caches stay alive)
        _currentPayrollRef = null;
        _cachedDeptRef = null;
        _cachedDeptName = "";

        dynamic conn = _sessionConn!;
        try
        {
            int created = 0;
            int skipped = 0;
            var errors = new List<string>();

            // Find our bank account ref (cached across files)
            dynamic? bankAcct = FindBankAccount(conn, file.Account);
            if (bankAcct == null)
            {
                return new LoadResult
                {
                    Success = false,
                    Error = $"Bank account {file.Account} not found in {mapping.Database}"
                };
            }

            // Get organization from bank account
            dynamic org = bankAcct.Владелец;

            // Department is cached by org name across files (persists in session)
            try
            {
                string orgName = (string)(org.Наименование ?? "");
                var dept = GetDeptByOrg(conn, orgName);
                _cachedDeptRef = dept.Ref;
                _cachedDeptName = dept.Name;
            }
            catch { }

            // Pre-load: ensure all counterparties and bank accounts exist
            EnsureCounterparties(conn, file);

            // Build duplicate counters: count existing docs in 1C per (date,amount,cpAcct,isDebit)
            Dictionary<string, int> existingCounts = CountExistingDocs(conn, file);
            _logger.LogInformation("  Existing documents in 1C: {Count} unique keys", existingCounts.Count);

            // --- Create shared ВедомостьНаВыплатуЗарплаты for salary (code 151) docs ---
            _currentPayrollRef = null;
            var salaryDocs = file.Documents
                .Where(d => d.IsDebit(file.Account)
                    && d.OperationType == "Перечисление заработной платы работнику")
                .ToList();

            if (salaryDocs.Count > 0)
            {
                try
                {
                    _logger.LogInformation("  Creating shared ВедомостьНаВыплатуЗарплаты for {Count} salary docs", salaryDocs.Count);
                    dynamic payroll = conn.Документы.ВедомостьНаВыплатуЗарплаты.СоздатьДокумент();

                    // Date from first salary doc
                    DateTime payrollDate = DateTime.ParseExact(salaryDocs[0].Date, "dd.MM.yyyy", CultureInfo.InvariantCulture);
                    payroll.Дата = payrollDate;
                    payroll.Организация = org;

                    // Total amount
                    decimal totalSalary = salaryDocs.Sum(d => d.Amount);
                    payroll.СуммаДокумента = totalSalary;

                    // ПодразделениеОрганизации (from cached lookup)
                    if (_cachedDeptRef != null)
                        try { payroll.ПодразделениеОрганизации = _cachedDeptRef; } catch { }

                    // ПериодРегистрации = start of month
                    payroll.ПериодРегистрации = new DateTime(payrollDate.Year, payrollDate.Month, 1);

                    // --- Enum fields via reflection ---
                    try
                    {
                        var val = GetEnumValue(conn, "ВидыМестВыплатыЗарплаты", "БанковскийСчет");
                        if (val != null) payroll.ВидМестаВыплаты = val;
                    }
                    catch (Exception ex) { _logger.LogWarning("  ВидМестаВыплаты failed: {0}", ex.Message); }

                    try
                    {
                        var val = GetEnumValue(conn, "СпособыРасчетовСФизическимиЛицами", "ОплатаТруда");
                        if (val != null) payroll.СпособРасчетов = val;
                    }
                    catch (Exception ex) { _logger.LogWarning("  СпособРасчетов failed: {0}", ex.Message); }

                    try
                    {
                        var val = GetEnumValue(conn, "ВидыДоходовИсполнительногоПроизводства", "ЗарплатаВознаграждения");
                        if (val != null) payroll.ВидДоходаИсполнительногоПроизводства = val;
                    }
                    catch (Exception ex) { _logger.LogWarning("  ВидДоходаИсполнительногоПроизводства failed: {0}", ex.Message); }

                    // --- Ensure ФизическиеЛица + bank accounts for salary ---
                    // For each salary doc:
                    // 1) Find bank account (of ФизическоеЛицо owner only)
                    //    → found: get ФизическоеЛицо from owner
                    // 2) Not found: search ФизическоеЛицо by name
                    //    → found: create bank account for this person
                    // 3) Not found by name: create ФизическоеЛицо + bank account
                    var salaryPhysMap = new List<(dynamic? PhysRef, BankDocument Doc)>();
                    foreach (var salDoc in salaryDocs)
                    {
                        string empName = ExtractPersonName(salDoc.RecipientName ?? "");
                        string empAcct = NormalizeAccount(salDoc.RecipientAccount ?? "");
                        dynamic? physRef = null;

                        // Step 1: find bank account owned by ФизическоеЛицо
                        if (!string.IsNullOrWhiteSpace(empAcct))
                        {
                            dynamic? bankAcctRef = FindBankAccountByNumber(conn, empAcct);
                            if (bankAcctRef != null)
                            {
                                try
                                {
                                    dynamic owner = bankAcctRef.Владелец;
                                    if (owner != null && !owner.Пустая())
                                    {
                                        // Owner could be Контрагент — find matching ФизическоеЛицо by name
                                        string ownerName = (string)(owner.Наименование ?? "");
                                        physRef = FindPhysicalPerson(conn, ownerName);
                                        if (physRef != null)
                                            _logger.LogInformation("    {Name}: found via bank account → {Found}",
                                                empName, ownerName);
                                    }
                                }
                                catch { }
                            }
                        }

                        // Step 2: search ФизическоеЛицо by name
                        if (physRef == null)
                        {
                            physRef = FindPhysicalPerson(conn, empName);
                            if (physRef != null)
                            {
                                string foundName = "";
                                try { foundName = (string)(physRef.Наименование ?? ""); } catch { }
                                _logger.LogInformation("    {Name}: found by name → {Found}", empName, foundName);

                                // Add new bank account — find matching Контрагент as owner
                                if (!string.IsNullOrWhiteSpace(empAcct))
                                {
                                    try
                                    {
                                        dynamic? cpRef = FindCounterpartyByName(conn, foundName);
                                        if (cpRef == null) cpRef = FindCounterpartyByName(conn, empName);
                                        if (cpRef != null)
                                            CreateBankAccountForOwner(conn, empAcct, cpRef, empName);
                                        else
                                            _logger.LogWarning("    No counterparty found to own bank account for {Name}", empName);
                                    }
                                    catch (Exception ex) { _logger.LogWarning("    Failed to create bank account: {Err}", ex.Message); }
                                }
                            }
                        }

                        // Step 3: create new ФизическоеЛицо + Контрагент + bank account
                        if (physRef == null)
                        {
                            _logger.LogInformation("    {Name}: creating new ФизическоеЛицо + Контрагент", empName);
                            try
                            {
                                dynamic newPhys = conn.Справочники.ФизическиеЛица.СоздатьЭлемент();
                                newPhys.Наименование = empName;
                                newPhys.Записать();
                                physRef = newPhys.Ссылка;

                                // Create Контрагент (ФизЛицо type) as bank account owner
                                dynamic newCp = conn.Справочники.Контрагенты.СоздатьЭлемент();
                                newCp.Наименование = empName;
                                try
                                {
                                    var val = GetEnumValue(conn, "ЮрФизЛицо", "ФизЛицо");
                                    if (val != null) newCp.ЮрФизЛицо = val;
                                }
                                catch { }
                                newCp.Записать();

                                if (!string.IsNullOrWhiteSpace(empAcct))
                                    CreateBankAccountForOwner(conn, empAcct, newCp.Ссылка, empName);
                            }
                            catch (Exception ex) { _logger.LogWarning("    Failed to create ФизическоеЛицо: {Err}", ex.Message); }
                        }

                        salaryPhysMap.Add((physRef, salDoc));
                    }

                    // --- Tabular section "Зарплата": one row per employee ---
                    dynamic tabSection = payroll.Зарплата;
                    foreach (var (physRef, salDoc) in salaryPhysMap)
                    {
                        dynamic newRow = tabSection.Добавить();
                        if (physRef != null)
                            newRow.ФизическоеЛицо = physRef;
                        newRow.Сумма = salDoc.Amount;
                    }

                    payroll.Записать();
                    _currentPayrollRef = payroll.Ссылка;
                    TrackCom(_currentPayrollRef);
                    _logger.LogInformation("  ВедомостьНаВыплатуЗарплаты created, {Count} rows, total {Total:F2}",
                        salaryDocs.Count, totalSalary);
                    // Release the document object itself (we only need the reference now)
                    SafeRelease(payroll);
                }
                catch (Exception ex)
                {
                    _logger.LogError("  Failed to create ВедомостьНаВыплатуЗарплаты: {Err}", ex.Message);
                    _currentPayrollRef = null;
                }
            }

            foreach (var doc in file.Documents)
            {
                try
                {
                    bool isDebit = doc.IsDebit(file.Account);
                    string cpAcct = isDebit ? doc.RecipientAccount : doc.PayerAccount;
                    // Key: date + amount + direction + doc number + counterparty account
                    string key = $"{doc.Date}|{doc.Amount.ToString("F2", CultureInfo.InvariantCulture)}|{(isDebit ? "D" : "C")}|{doc.Number}|{NormalizeAccount(cpAcct)}";
                    _logger.LogDebug("  File doc key={Key}, existing count={Count}", key, existingCounts.GetValueOrDefault(key));

                    if (existingCounts.TryGetValue(key, out int remaining) && remaining > 0)
                    {
                        existingCounts[key] = remaining - 1;
                        _logger.LogInformation("  SKIP (exists): {0} {1} {2}",
                            doc.Date, doc.Amount, NormalizeAccount(cpAcct));
                        skipped++;
                        continue;
                    }

                    bool wasCreated = CreateDocument(conn, doc, bankAcct, org, isDebit, file.Account);
                    if (wasCreated)
                        created++;
                    else
                        skipped++;
                }
                catch (Exception ex)
                {
                    errors.Add($"Doc #{doc.Number}: {ex.Message}");
                    skipped++;
                }
            }

            return new LoadResult
            {
                Success = errors.Count == 0,
                Created = created,
                Skipped = skipped,
                Error = errors.Count > 0
                    ? string.Join("; ", errors)
                    : null
            };
        }
        catch (Exception ex)
        {
            // On COM error the session may be broken (e.g. 1C killed it). Tear it down so
            // the next file gets a fresh connection.
            try { EndSession(); } catch { }
            return new LoadResult
            {
                Success = false,
                Error = $"COM error: {ex.Message}"
            };
        }
        finally
        {
            // Per-file teardown: drop transient refs but keep session caches alive.
            _currentPayrollRef = null;
            // _cachedDeptRef stays in _cacheDeptByOrg; just clear the per-file shortcut
            _cachedDeptRef = null;
            _cachedDeptName = "";
        }
    }

    private dynamic? FindBankAccount(dynamic conn, string account)
    {
        string normalized = NormalizeAccount(account);
        if (_cacheBankAcctsByNumber.TryGetValue(normalized, out var cached) && cached != null)
            return cached;
        // Try without IBAN prefix, then with ME25 prefix
        foreach (var variant in new[] { normalized, "ME25" + normalized })
        {
            try
            {
                dynamic found = conn.Справочники.БанковскиеСчета
                    .НайтиПоРеквизиту("НомерСчета", variant);
                if (!found.Пустая() && !(bool)found.ПометкаУдаления)
                {
                    _cacheBankAcctsByNumber[normalized] = found;
                    TrackCom(found);
                    return found;
                }
            }
            catch { }
        }
        return null;
    }

    private dynamic? FindBankAccountByNumber(dynamic conn, string account)
    {
        string normalized = NormalizeAccount(account);
        if (string.IsNullOrEmpty(normalized)) return null;
        if (_cacheBankAcctsByNumber.TryGetValue(normalized, out var cached) && cached != null)
            return cached;

        foreach (var variant in new[] { normalized, "ME25" + normalized })
        {
            try
            {
                dynamic found = conn.Справочники.БанковскиеСчета
                    .НайтиПоРеквизиту("НомерСчета", variant);
                if (!found.Пустая() && !(bool)found.ПометкаУдаления)
                {
                    _cacheBankAcctsByNumber[normalized] = found;
                    TrackCom(found);
                    return found;
                }
            }
            catch { }
        }
        // Don't cache misses — bank accounts may be created during EnsureCounterparties
        return null;
    }

    /// <summary>Extract person name: strip address after comma, keep first 2 words.</summary>
    private static string ExtractPersonName(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        // Strip everything after first comma (address)
        int comma = raw.IndexOf(',');
        if (comma > 0) raw = raw[..comma];
        // Keep only first 2 words (first name + last name)
        var words = raw.Trim().Split(' ', StringSplitOptions.RemoveEmptyEntries);
        return words.Length <= 2
            ? string.Join(" ", words)
            : $"{words[0]} {words[1]}";
    }

    public static string NormalizeAccount(string acct)
    {
        if (string.IsNullOrEmpty(acct)) return "";
        // Strip IBAN prefix (ME25), dashes, spaces
        var s = acct.Replace("-", "").Replace(" ", "");
        if (s.StartsWith("ME") && s.Length > 4 && char.IsDigit(s[2]))
            s = s[4..];
        return s;
    }

    /// <summary>Count existing documents in 1C grouped by (docNum,date,amount,cpAcct,direction).</summary>
    private Dictionary<string, int> CountExistingDocs(dynamic conn, BankExchangeFile file)
    {
        var counts = new Dictionary<string, int>();
        try
        {
            var dates = file.Documents
                .Select(d => DateTime.ParseExact(d.Date, "dd.MM.yyyy", CultureInfo.InvariantCulture))
                .ToList();
            if (dates.Count == 0) return counts;
            DateTime minDate = dates.Min().Date;
            DateTime maxDate = dates.Max().Date.AddDays(1).AddSeconds(-1);

            // Use queries instead of selection to avoid COM memory leaks
            foreach (var (docType, dir) in new[] { ("СписаниеСРасчетногоСчета", "D"), ("ПоступлениеНаРасчетныйСчет", "C") })
            {
                dynamic? query = null;
                dynamic? result = null;
                dynamic? sel = null;
                try
                {
                    query = conn.NewObject("Запрос");
                    query.Текст = $@"ВЫБРАТЬ
                        Док.СуммаДокумента КАК Сумма,
                        Док.Дата КАК Дата,
                        Док.НомерВходящегоДокумента КАК НомерВх,
                        Док.СчетКонтрагента.НомерСчета КАК СчетКП
                    ИЗ
                        Документ.{docType} КАК Док
                    ГДЕ
                        Док.Дата МЕЖДУ &Дата1 И &Дата2
                        И НЕ Док.ПометкаУдаления";
                    query.УстановитьПараметр("Дата1", minDate);
                    query.УстановитьПараметр("Дата2", maxDate);

                    result = query.Выполнить();
                    sel = result.Выбрать();
                    while (sel.Следующий())
                    {
                        try
                        {
                            decimal amt = (decimal)sel.Сумма;
                            DateTime dt = ((DateTime)sel.Дата).Date;
                            string docNum = (string)(sel.НомерВх ?? "");
                            string cpAcct = NormalizeAccount((string)(sel.СчетКП ?? ""));
                            string key = $"{dt:dd.MM.yyyy}|{amt.ToString("F2", CultureInfo.InvariantCulture)}|{dir}|{docNum}|{cpAcct}";
                            counts[key] = counts.GetValueOrDefault(key) + 1;
                            _logger.LogDebug("    1C existing: key={Key}", key);
                        }
                        catch { }
                    }
                }
                catch { }
                finally
                {
                    SafeRelease(sel);
                    SafeRelease(result);
                    SafeRelease(query);
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning("  CountExistingDocs failed: {0}", ex.Message);
        }
        return counts;
    }

    /// <summary>
    /// Find ФизическоеЛицо by name with word-order swap and case-insensitive prefix search.
    /// Tries: "Firstname Lastname", "Lastname Firstname" as prefixes.
    /// </summary>
    public dynamic? FindPhysicalPerson(dynamic conn, string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;

        // Generate name variants: original + swapped word order
        var variants = new List<string> { name };
        var words = name.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 2)
            variants.Add($"{words[1]} {words[0]}");
        else if (words.Length == 3)
        {
            // Could be "Firstname Middle Lastname" etc.
            variants.Add($"{words[2]} {words[0]} {words[1]}");
            variants.Add($"{words[1]} {words[2]} {words[0]}");
        }

        foreach (var variant in variants)
        {
            try
            {
                dynamic found = conn.Справочники.ФизическиеЛица
                    .НайтиПоНаименованию(variant, false);
                if (!found.Пустая() && !(bool)found.ПометкаУдаления)
                {
                    // Verify it's a ФизическоеЛицо element, not a group or wrong type
                    try
                    {
                        bool isGroup = false;
                        try { isGroup = (bool)found.ЭтоГруппа; } catch { }
                        if (isGroup) continue;

                        // Verify name actually starts with our search (НайтиПоНаименованию false can return unexpected results)
                        string foundName = ((string)(found.Наименование ?? "")).Trim();
                        if (foundName.StartsWith(variant, StringComparison.OrdinalIgnoreCase))
                            return found;
                    }
                    catch { }
                }
            }
            catch { }
        }
        return null;
    }

    /// <summary>
    /// Find Контрагент by name: exact → prefix → stripped suffix → word swap.
    /// </summary>
    public dynamic? FindCounterpartyByName(dynamic conn, string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        if (_cacheCounterpartyByName.TryGetValue(name, out var cached) && cached != null)
            return cached;

        // Try exact first
        try
        {
            dynamic found = conn.Справочники.Контрагенты
                .НайтиПоНаименованию(name, true);
            if (!found.Пустая() && !(bool)found.ПометкаУдаления)
            {
                _cacheCounterpartyByName[name] = found;
                TrackCom(found);
                return found;
            }
        }
        catch { }

        // Try prefix with word-order swap + stripped suffix
        var variants = new List<string> { name };
        var words = name.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (words.Length == 2)
            variants.Add($"{words[1]} {words[0]}");
        // Add stripped variant (remove DOO/AD/BANKA etc.)
        string stripped = StripCompanySuffix(name);
        if (!string.IsNullOrWhiteSpace(stripped) && stripped != name)
            variants.Add(stripped);

        foreach (var variant in variants)
        {
            try
            {
                dynamic found = conn.Справочники.Контрагенты
                    .НайтиПоНаименованию(variant, false);
                if (!found.Пустая() && !(bool)found.ПометкаУдаления)
                {
                    _cacheCounterpartyByName[name] = found;
                    TrackCom(found);
                    return found;
                }
            }
            catch { }
        }
        // Don't cache misses — may be created later
        return null;
    }

    public void Dispose() { }
}
