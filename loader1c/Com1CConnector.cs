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
    private readonly string _mappingFile;
    private readonly string _reportFile;

    // Shared payroll document reference for current file being loaded
    private dynamic? _currentPayrollRef;
    private dynamic? _cachedDeptRef;
    private string _cachedDeptName = "";

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
    }

    public IReadOnlyDictionary<string, AccountMapping> AccountMap => _accountMap;

    /// <summary>Connect to a specific database.</summary>
    private dynamic Connect(string database)
    {
        var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
            ?? throw new Exception("V83.COMConnector not registered");
        dynamic connector = Activator.CreateInstance(connectorType)!;
        var connStr = $"Srvr=\"{_config.Server}\";Ref=\"{database}\";" +
                      $"Usr=\"{_config.User}\";Pwd=\"{_config.Password}\"";
        return connector.Connect(connStr);
    }

    /// <summary>Scan all databases and build account→database mapping.</summary>
    public void ScanDatabases()
    {
        _logger.LogInformation("Scanning {Count} databases...", _config.Databases.Count);
        var map = new Dictionary<string, AccountMapping>();

        foreach (var db in _config.Databases)
        {
            try
            {
                _logger.LogInformation("Scanning {Database}...", db);
                dynamic conn = Connect(db);

                // Read bank accounts with owner organization
                dynamic acctSel = conn.Справочники.БанковскиеСчета.Выбрать();
                while (acctSel.Следующий())
                {
                    try
                    {
                        bool deleted = acctSel.ПометкаУдаления;
                        if (deleted) continue;

                        string acctNo = acctSel.НомерСчета ?? "";
                        if (string.IsNullOrWhiteSpace(acctNo)) continue;

                        // Normalize: strip IBAN prefix, dashes, spaces
                        string normalized = NormalizeAccount(acctNo);
                        if (normalized.Length < 10) continue;

                        string orgName = "";
                        string inn = "";
                        try
                        {
                            orgName = acctSel.Владелец.Наименование ?? "";
                            inn = acctSel.Владелец.ИНН ?? "";
                        }
                        catch { /* owner might not be an organization */ }

                        if (!map.ContainsKey(normalized))
                        {
                            map[normalized] = new AccountMapping
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

                Marshal.ReleaseComObject(conn);
            }
            catch (Exception ex)
            {
                _logger.LogError("Failed to scan {Database}: {Error}", db, ex.Message);
            }
        }

        _accountMap = map;
        _logger.LogInformation("Scan complete: {Count} accounts mapped", map.Count);

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

    /// <summary>Load mapping from cached file.</summary>
    public bool LoadCachedMapping()
    {
        if (!File.Exists(_mappingFile)) return false;
        try
        {
            var json = File.ReadAllText(_mappingFile, Encoding.UTF8);
            _accountMap = JsonSerializer.Deserialize<Dictionary<string, AccountMapping>>(json)
                          ?? new();
            _logger.LogInformation("Loaded cached mapping: {Count} accounts", _accountMap.Count);
            return _accountMap.Count > 0;
        }
        catch
        {
            return false;
        }
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

        try
        {
            dynamic conn = Connect(mapping.Database);
            int created = 0;
            int skipped = 0;
            var errors = new List<string>();

            // Find our bank account ref
            dynamic bankAcct = FindBankAccount(conn, file.Account);
            if (bankAcct == null)
            {
                Marshal.ReleaseComObject(conn);
                return new LoadResult
                {
                    Success = false,
                    Error = $"Bank account {file.Account} not found in {mapping.Database}"
                };
            }

            // Get organization from bank account
            dynamic org = bankAcct.Владелец;

            // Find department once for this organization
            _cachedDeptRef = null;
            _cachedDeptName = "";
            try
            {
                string orgName = (string)(org.Наименование ?? "");
                dynamic deptSel = conn.Справочники.ПодразделенияОрганизаций.Выбрать();
                while (deptSel.Следующий())
                {
                    if ((bool)deptSel.ПометкаУдаления) continue;
                    try
                    {
                        string ownerName = (string)(deptSel.Владелец?.Наименование ?? "");
                        if (ownerName == orgName)
                        {
                            _cachedDeptRef = deptSel.Ссылка;
                            _cachedDeptName = (string)(deptSel.Наименование ?? "");
                            break;
                        }
                    }
                    catch { }
                }
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
                    // ВидМестаВыплаты = БанковскийСчет
                    try
                    {
                        dynamic enumMgr = conn.Перечисления.ВидыМестВыплатыЗарплаты;
                        var val = enumMgr.GetType().InvokeMember("БанковскийСчет",
                            System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                        payroll.ВидМестаВыплаты = val;
                    }
                    catch (Exception ex) { _logger.LogWarning("  ВидМестаВыплаты failed: {0}", ex.Message); }

                    // СпособРасчетов = ОплатаТруда
                    try
                    {
                        dynamic enumMgr = conn.Перечисления.СпособыРасчетовСФизическимиЛицами;
                        var val = enumMgr.GetType().InvokeMember("ОплатаТруда",
                            System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                        payroll.СпособРасчетов = val;
                    }
                    catch (Exception ex) { _logger.LogWarning("  СпособРасчетов failed: {0}", ex.Message); }

                    // ВидДоходаИсполнительногоПроизводства = ЗарплатаВознаграждения
                    try
                    {
                        dynamic enumMgr = conn.Перечисления.ВидыДоходовИсполнительногоПроизводства;
                        var val = enumMgr.GetType().InvokeMember("ЗарплатаВознаграждения",
                            System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                        payroll.ВидДоходаИсполнительногоПроизводства = val;
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
                        string empName = salDoc.RecipientName ?? "";
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
                                    dynamic enumMgr = conn.Перечисления.ЮрФизЛицо;
                                    var val = enumMgr.GetType().InvokeMember("ФизЛицо",
                                        System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                                    newCp.ЮрФизЛицо = val;
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
                    _logger.LogInformation("  ВедомостьНаВыплатуЗарплаты created, {Count} rows, total {Total:F2}",
                        salaryDocs.Count, totalSalary);
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
                    // Key: doc date + amount + direction (with counting for identical ops)
                    string key = $"{doc.Date}|{doc.Amount.ToString("F2", CultureInfo.InvariantCulture)}|{(isDebit ? "D" : "C")}";
                    _logger.LogInformation("  File doc key={Key}, existing count={Count}", key, existingCounts.GetValueOrDefault(key));

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

            Marshal.ReleaseComObject(conn);

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
            return new LoadResult
            {
                Success = false,
                Error = $"COM error: {ex.Message}"
            };
        }
    }

    private dynamic? FindBankAccount(dynamic conn, string account)
    {
        string normalized = NormalizeAccount(account);
        // Try without IBAN prefix, then with ME25 prefix
        foreach (var variant in new[] { normalized, "ME25" + normalized })
        {
            try
            {
                dynamic found = conn.Справочники.БанковскиеСчета
                    .НайтиПоРеквизиту("НомерСчета", variant);
                if (!found.Пустая() && !(bool)found.ПометкаУдаления)
                    return found;
            }
            catch { }
        }
        return null;
    }

    private dynamic? FindBankAccountByNumber(dynamic conn, string account)
    {
        string normalized = NormalizeAccount(account);
        foreach (var variant in new[] { normalized, "ME25" + normalized })
        {
            try
            {
                dynamic found = conn.Справочники.БанковскиеСчета
                    .НайтиПоРеквизиту("НомерСчета", variant);
                if (!found.Пустая() && !(bool)found.ПометкаУдаления)
                    return found;
            }
            catch { }
        }
        return null;
    }

    private static string NormalizeAccount(string acct)
    {
        if (string.IsNullOrEmpty(acct)) return "";
        // Strip IBAN prefix (ME25), dashes, spaces
        var s = acct.Replace("-", "").Replace(" ", "");
        if (s.StartsWith("ME") && s.Length > 4 && char.IsDigit(s[2]))
            s = s[4..];
        return s;
    }

    /// <summary>Count existing documents in 1C grouped by (stmtNum,stmtDate,amount,cpAcct,direction).</summary>
    private Dictionary<string, int> CountExistingDocs(dynamic conn, BankExchangeFile file)
    {
        var counts = new Dictionary<string, int>();
        try
        {
            // Get date range from file
            var dates = file.Documents
                .Select(d => DateTime.ParseExact(d.Date, "dd.MM.yyyy", CultureInfo.InvariantCulture))
                .ToList();
            if (dates.Count == 0) return counts;
            DateTime minDate = dates.Min().Date;
            DateTime maxDate = dates.Max().Date.AddDays(1).AddSeconds(-1);

            foreach (var docType in new[] { ("СписаниеСРасчетногоСчета", "D"), ("ПоступлениеНаРасчетныйСчет", "C") })
            {
                try
                {
                    dynamic mgr = conn.Документы.GetType().InvokeMember(
                        docType.Item1, System.Reflection.BindingFlags.GetProperty,
                        null, (object)conn.Документы, null);
                    dynamic selection = mgr.Выбрать(minDate, maxDate);

                    while (selection.Следующий())
                    {
                        try
                        {
                            if ((bool)selection.ПометкаУдаления) continue;

                            decimal amt = (decimal)selection.СуммаДокумента;
                            DateTime dt = ((DateTime)selection.Дата).Date;
                            string key = $"{dt:dd.MM.yyyy}|{amt.ToString("F2", CultureInfo.InvariantCulture)}|{docType.Item2}";
                            counts[key] = counts.GetValueOrDefault(key) + 1;
                            _logger.LogInformation("    1C existing: key={Key}", key);
                        }
                        catch { }
                    }
                }
                catch { }
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

        // Try exact first
        try
        {
            dynamic found = conn.Справочники.Контрагенты
                .НайтиПоНаименованию(name, true);
            if (!found.Пустая() && !(bool)found.ПометкаУдаления)
                return found;
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
                    return found;
            }
            catch { }
        }
        return null;
    }

    public void Dispose() { }
}
