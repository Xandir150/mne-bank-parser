using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

// ──────────────────────────────────────────────────────────────
// Configuration
// ──────────────────────────────────────────────────────────────
public class LoaderConfig
{
    public string Server { get; set; } = "navus-server";
    public string User { get; set; } = "";
    public string Password { get; set; } = "";
    public string OutputDir { get; set; } = "";
    public string LoadedDir { get; set; } = "";
    public string LogFile { get; set; } = "";
    public int ScanIntervalSec { get; set; } = 30;
    public bool PostDocuments { get; set; } = false;
    public List<string> Databases { get; set; } = new();
}

// ──────────────────────────────────────────────────────────────
// Account mapping: bank account number → (database, orgName)
// ──────────────────────────────────────────────────────────────
public class AccountMapping
{
    public string Database { get; set; } = "";
    public string OrgName { get; set; } = "";
    public string Inn { get; set; } = "";
}

// ──────────────────────────────────────────────────────────────
// Parsed bank exchange file
// ──────────────────────────────────────────────────────────────
public class BankExchangeFile
{
    public string FilePath { get; set; } = "";
    public string Account { get; set; } = "";
    public string DateStart { get; set; } = "";
    public string DateEnd { get; set; } = "";
    public List<BankDocument> Documents { get; set; } = new();
}

public class BankDocument
{
    public string Number { get; set; } = "";
    public string Date { get; set; } = "";
    public decimal Amount { get; set; }
    public string PayerAccount { get; set; } = "";
    public string PayerName { get; set; } = "";
    public string PayerInn { get; set; } = "";
    public string RecipientAccount { get; set; } = "";
    public string RecipientName { get; set; } = "";
    public string RecipientInn { get; set; } = "";
    public string Purpose { get; set; } = "";

    // Extended fields from export_1c.py
    public string OperationType { get; set; } = "";       // ВидОперации
    public string DebitAccount { get; set; } = "";         // СчетДебета
    public string CreditAccount { get; set; } = "";        // СчетУчета (51)
    public string CashFlowItem { get; set; } = "";         // СтатьяДвиженияДенежныхСредств
    public string TaxType { get; set; } = "";              // ВидНалога
    public string ObligationType { get; set; } = "";       // ВидОбязательства
    public string ExpenseItem { get; set; } = "";          // СтатьяРасходов
    public string StatementNumber { get; set; } = "";      // НомерВыписки
    public string StatementDate { get; set; } = "";        // ДатаВыписки

    /// <summary>True if money leaves our account (debit).</summary>
    public bool IsDebit(string ourAccount)
    {
        return NormalizeAccount(PayerAccount) == NormalizeAccount(ourAccount);
    }

    static string NormalizeAccount(string a) =>
        (a ?? "").Replace("-", "").Replace(" ", "").TrimStart('0');
}

// ──────────────────────────────────────────────────────────────
// File parser for 1CClientBankExchange format
// ──────────────────────────────────────────────────────────────
public static class BankFileParser
{
    public static BankExchangeFile Parse(string filePath)
    {
        var result = new BankExchangeFile { FilePath = filePath };
        var lines = File.ReadAllLines(filePath, Encoding.GetEncoding(1251));

        BankDocument? currentDoc = null;

        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (string.IsNullOrEmpty(trimmed)) continue;

            if (trimmed.StartsWith("РасчСчет="))
                result.Account = trimmed[9..];
            else if (trimmed.StartsWith("ДатаНачала="))
                result.DateStart = trimmed[11..];
            else if (trimmed.StartsWith("ДатаКонца="))
                result.DateEnd = trimmed[10..];
            else if (trimmed.StartsWith("СекцияДокумент="))
            {
                currentDoc = new BankDocument();
            }
            else if (trimmed == "КонецДокумента")
            {
                if (currentDoc != null)
                    result.Documents.Add(currentDoc);
                currentDoc = null;
            }
            else if (currentDoc != null)
            {
                if (trimmed.StartsWith("Номер="))
                    currentDoc.Number = trimmed[6..];
                else if (trimmed.StartsWith("Дата="))
                    currentDoc.Date = trimmed[5..];
                else if (trimmed.StartsWith("Сумма="))
                {
                    if (decimal.TryParse(trimmed[6..], NumberStyles.Any,
                            CultureInfo.InvariantCulture, out var a))
                        currentDoc.Amount = a;
                }
                else if (trimmed.StartsWith("ПлательщикСчет="))
                    currentDoc.PayerAccount = trimmed[15..];
                else if (trimmed.StartsWith("Плательщик="))
                    currentDoc.PayerName = trimmed[11..];
                else if (trimmed.StartsWith("ПлательщикИНН="))
                    currentDoc.PayerInn = trimmed[14..];
                else if (trimmed.StartsWith("ПолучательСчет="))
                    currentDoc.RecipientAccount = trimmed[15..];
                else if (trimmed.StartsWith("Получатель="))
                    currentDoc.RecipientName = trimmed[11..];
                else if (trimmed.StartsWith("ПолучательИНН="))
                    currentDoc.RecipientInn = trimmed[14..];
                else if (trimmed.StartsWith("НазначениеПлатежа="))
                    currentDoc.Purpose = trimmed[18..];
                    // Extended fields
                else if (trimmed.StartsWith("ВидОперации="))
                    currentDoc.OperationType = trimmed[12..];
                else if (trimmed.StartsWith("СчетДебета="))
                    currentDoc.DebitAccount = trimmed[11..];
                else if (trimmed.StartsWith("СчетУчета="))
                    currentDoc.CreditAccount = trimmed[10..];
                else if (trimmed.StartsWith("СтатьяДвиженияДенежныхСредств="))
                    currentDoc.CashFlowItem = trimmed[30..];
                else if (trimmed.StartsWith("ВидНалога="))
                    currentDoc.TaxType = trimmed[10..];
                else if (trimmed.StartsWith("ВидОбязательства="))
                    currentDoc.ObligationType = trimmed[17..];
                else if (trimmed.StartsWith("СтатьяРасходов="))
                    currentDoc.ExpenseItem = trimmed[15..];
                else if (trimmed.StartsWith("НомерВыписки="))
                    currentDoc.StatementNumber = trimmed[13..];
                else if (trimmed.StartsWith("ДатаВыписки="))
                    currentDoc.StatementDate = trimmed[12..];
            }
        }

        return result;
    }
}

// ──────────────────────────────────────────────────────────────
// 1C COM connector — all COM calls run on a dedicated STA thread
// ──────────────────────────────────────────────────────────────
public class Com1CConnector : IDisposable
{
    private readonly LoaderConfig _config;
    private readonly ILogger _logger;

    // Account number → mapping
    private Dictionary<string, AccountMapping> _accountMap = new();
    private readonly string _mappingFile;
    private readonly string _reportFile;

    // Shared payroll document reference for current file being loaded
    private dynamic? _currentPayrollRef;

    public Com1CConnector(LoaderConfig config, ILogger logger)
    {
        _config = config;
        _logger = logger;
        var baseDir = AppContext.BaseDirectory;
        _mappingFile = Path.Combine(baseDir, "account_mapping.json");
        _reportFile = Path.Combine(baseDir, "account_mapping.txt");
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

    /// <summary>Find which database an account belongs to.</summary>
    public AccountMapping? FindAccount(string account)
    {
        var normalized = NormalizeAccount(account);
        if (_accountMap.TryGetValue(normalized, out var mapping))
            return mapping;

        // Try partial match (some accounts might have extra digits)
        foreach (var kvp in _accountMap)
        {
            if (kvp.Key.Contains(normalized) || normalized.Contains(kvp.Key))
                return kvp.Value;
        }

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

                    // ПодразделениеОрганизации
                    try
                    {
                        dynamic deptSel = conn.Справочники.ПодразделенияОрганизаций.Выбрать();
                        while (deptSel.Следующий())
                        {
                            if ((bool)deptSel.ПометкаУдаления) continue;
                            try
                            {
                                dynamic owner = deptSel.Владелец;
                                string ownerName = (string)(owner?.Наименование ?? "");
                                string orgName = (string)(org.Наименование ?? "");
                                if (ownerName == orgName)
                                {
                                    payroll.ПодразделениеОрганизации = deptSel.Ссылка;
                                    break;
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }

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

                    // --- Tabular section "Зарплата": one row per employee ---
                    dynamic tabSection = payroll.Зарплата;
                    foreach (var salDoc in salaryDocs)
                    {
                        // Find ФизическоеЛицо via counterparty name (recipient)
                        string empName = salDoc.RecipientName ?? "";
                        dynamic? empRef = null;
                        if (!string.IsNullOrWhiteSpace(empName))
                        {
                            try
                            {
                                empRef = conn.Справочники.ФизическиеЛица
                                    .НайтиПоНаименованию(empName, true);
                                if (empRef.Пустая()) empRef = null;
                            }
                            catch { }
                        }

                        dynamic newRow = tabSection.Добавить();
                        if (empRef != null)
                            newRow.ФизическоеЛицо = empRef;
                        newRow.Сумма = salDoc.Amount;
                        _logger.LogInformation("    Ведомость row: {Name} = {Amount:F2}", empName, salDoc.Amount);
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

                    bool wasCreated = CreateDocument(conn, doc, bankAcct, org, isDebit);
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
        dynamic sel = conn.Справочники.БанковскиеСчета.Выбрать();
        while (sel.Следующий())
        {
            try
            {
                string no = NormalizeAccount(sel.НомерСчета ?? "");
                if (no == normalized)
                    return sel.Ссылка;
            }
            catch { }
        }
        return null;
    }

    // ВидОперации mapping: our Russian name → 1C enum value name
    // Enums: ВидыОперацийСписаниеДенежныхСредств / ВидыОперацийПоступлениеДенежныхСредств
    private static readonly Dictionary<string, (string DebitEnum, string CreditEnum)> _opTypeMap = new()
    {
        ["Уплата налога"]       = ("ПеречислениеНалога", "ВозвратНалога"),
        ["Комиссия банка"]      = ("КомиссияБанка", "ПрочееПоступление"),
        ["Оплата поставщику"]   = ("ОплатаПоставщику", "ОплатаПокупателя"),
        ["Оплата от покупателя"]= ("ОплатаПоставщику", "ОплатаПокупателя"),
        ["Прочее списание"]     = ("ПрочееСписание", "ПрочееПоступление"),
        ["Прочее поступление"]  = ("ПрочееСписание", "ПрочееПоступление"),
        ["Перечисление заработной платы работнику"] = ("ПеречислениеЗаработнойПлатыРаботнику", "ПрочееПоступление"),
        ["Перечисление заработной платы по ведомостям"] = ("ПеречислениеЗП", "ПрочееПоступление"),
    };

    /// <summary>Returns true if document was created.</summary>
    private bool CreateDocument(dynamic conn, BankDocument doc,
        dynamic bankAcct, dynamic org, bool isDebit)
    {
        _logger.LogInformation("  >> Doc: {Amount} {Op} {Purpose}",
            doc.Amount, doc.OperationType, doc.Purpose?.Substring(0, Math.Min(40, doc.Purpose?.Length ?? 0)));

        // Parse date
        DateTime docDate = DateTime.ParseExact(doc.Date, "dd.MM.yyyy",
            CultureInfo.InvariantCulture);

        dynamic newDoc;
        if (isDebit)
        {
            newDoc = conn.Документы.СписаниеСРасчетногоСчета.СоздатьДокумент();
        }
        else
        {
            newDoc = conn.Документы.ПоступлениеНаРасчетныйСчет.СоздатьДокумент();
        }

        newDoc.Дата = docDate;
        newDoc.Организация = org;
        // Don't set Номер — let 1C auto-assign to avoid NullReferenceException

        // Bank account of our organization
        try { newDoc.СчетОрганизации = bankAcct; } catch { }

        // СчетБанк — the accounting account (план счетов), e.g. "51"
        string acctCode = "51";
        try
        {
            dynamic chartAcct = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду(acctCode);
            if (!chartAcct.Пустая())
                newDoc.СчетБанк = chartAcct;
            else
                _logger.LogWarning("    Chart account '{Code}' not found", acctCode);
        }
        catch (Exception ex)
        {
            _logger.LogWarning("    Failed to set СчетБанк '{Code}': {Err}", acctCode, ex.Message);
        }

        // ВалютаДокумента = EUR
        try
        {
            dynamic eur = conn.Справочники.Валюты.НайтиПоКоду("978");
            if (eur.Пустая())
                eur = conn.Справочники.Валюты.НайтиПоНаименованию("EUR", true);
            if (!eur.Пустая())
                newDoc.ВалютаДокумента = eur;
            else
                _logger.LogWarning("    Currency EUR not found");
        }
        catch (Exception ex)
        {
            _logger.LogWarning("    Failed to set ВалютаДокумента: {0}", ex.Message);
        }

        // НомерВходящегоДокумента / ДатаВходящегоДокумента
        if (!string.IsNullOrWhiteSpace(doc.StatementNumber))
        {
            try { newDoc.НомерВходящегоДокумента = doc.StatementNumber; } catch { }
        }
        if (!string.IsNullOrWhiteSpace(doc.StatementDate))
        {
            try
            {
                DateTime stmtDate = DateTime.ParseExact(doc.StatementDate, "dd.MM.yyyy",
                    CultureInfo.InvariantCulture);
                newDoc.ДатаВходящегоДокумента = stmtDate;
            }
            catch { }
        }

        // Amount
        newDoc.СуммаДокумента = doc.Amount;

        // --- ВидОперации ---
        if (!string.IsNullOrWhiteSpace(doc.OperationType))
        {
            try
            {
                string enumName;
                if (_opTypeMap.TryGetValue(doc.OperationType, out var mapped))
                    enumName = isDebit ? mapped.DebitEnum : mapped.CreditEnum;
                else
                    enumName = isDebit ? "ПрочееСписание" : "ПрочееПоступление";

                // COM enum access: use reflection since dynamic indexer doesn't work
                if (isDebit)
                {
                    dynamic enumMgr = conn.Перечисления.ВидыОперацийСписаниеДенежныхСредств;
                    var enumVal = enumMgr.GetType().InvokeMember(
                        enumName, System.Reflection.BindingFlags.GetProperty,
                        null, (object)enumMgr, null);
                    newDoc.ВидОперации = enumVal;
                }
                else
                {
                    dynamic enumMgr = conn.Перечисления.ВидыОперацийПоступлениеДенежныхСредств;
                    var enumVal = enumMgr.GetType().InvokeMember(
                        enumName, System.Reflection.BindingFlags.GetProperty,
                        null, (object)enumMgr, null);
                    newDoc.ВидОперации = enumVal;
                }

                _logger.LogInformation("    ВидОперации: {Op} → {Enum}",
                    doc.OperationType, enumName);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("    Failed to set ВидОперации '{Op}': {Err}",
                    doc.OperationType, ex.Message);
            }
        }

        // --- Counterparty ---
        string counterpartyName = isDebit ? doc.RecipientName : doc.PayerName;
        string counterpartyInn = isDebit ? doc.RecipientInn : doc.PayerInn;
        string counterpartyAcct = isDebit ? doc.RecipientAccount : doc.PayerAccount;

        // Find counterparty bank account first (needed for both СчетКонтрагента and owner lookup)
        dynamic cpBankAcctRef = null;
        if (!string.IsNullOrWhiteSpace(counterpartyAcct))
        {
            try
            {
                cpBankAcctRef = FindBankAccountByNumber(conn, counterpartyAcct);
                if (cpBankAcctRef != null)
                {
                    try { newDoc.СчетКонтрагента = cpBankAcctRef; } catch { }
                }
            }
            catch { }
        }

        // Search counterparty: by INN → by name → by bank account owner
        bool counterpartySet = false;
        if (!string.IsNullOrWhiteSpace(counterpartyInn))
        {
            try
            {
                dynamic found = conn.Справочники.Контрагенты
                    .НайтиПоРеквизиту("ИНН", counterpartyInn);
                if (!found.Пустая())
                {
                    newDoc.Контрагент = found;
                    counterpartySet = true;
                }
            }
            catch { }
        }
        if (!counterpartySet && !string.IsNullOrWhiteSpace(counterpartyName))
        {
            try
            {
                dynamic found = conn.Справочники.Контрагенты
                    .НайтиПоНаименованию(counterpartyName, true);
                if (!found.Пустая())
                {
                    newDoc.Контрагент = found;
                    counterpartySet = true;
                }
            }
            catch { }
        }
        // Fallback: find counterparty via bank account owner
        if (!counterpartySet && cpBankAcctRef != null)
        {
            try
            {
                dynamic owner = cpBankAcctRef.Владелец;
                if (owner != null && !owner.Пустая())
                {
                    newDoc.Контрагент = owner;
                    counterpartySet = true;
                    _logger.LogInformation("    Counterparty from bank account owner: {Name}",
                        (string)(owner.Наименование ?? ""));
                }
            }
            catch { }
        }

        // Note: СчетДебета/СчетКредита are NOT document properties in 1C:Бухгалтерия.
        // Accounting accounts are determined by the operation type during posting.
        // We set ВидОперации which controls which accounts are used.

        // --- СтатьяДвиженияДенежныхСредств ---
        if (!string.IsNullOrWhiteSpace(doc.CashFlowItem))
        {
            try
            {
                dynamic item = conn.Справочники
                    .СтатьиДвиженияДенежныхСредств
                    .НайтиПоНаименованию(doc.CashFlowItem, true);
                if (!item.Пустая())
                    newDoc.СтатьяДвиженияДенежныхСредств = item;
            }
            catch (Exception ex)
            {
                _logger.LogWarning("    Failed to set СтатьяДДС '{Item}': {Err}",
                    doc.CashFlowItem, ex.Message);
            }
        }

        // --- Tax fields (for ПеречислениеНалога) ---
        // Document property: Налог (ref to catalog ВидыНалоговИПлатежейВБюджет)
        if (!string.IsNullOrWhiteSpace(doc.TaxType))
        {
            try
            {
                dynamic taxRef = conn.Справочники.ВидыНалоговИПлатежейВБюджет
                    .НайтиПоНаименованию(doc.TaxType, true);
                if (!taxRef.Пустая())
                    newDoc.Налог = taxRef;
                else
                    _logger.LogWarning("    Tax type not found: '{Tax}'", doc.TaxType);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("    Failed to set Налог '{Tax}': {Err}",
                    doc.TaxType, ex.Message);
            }
        }

        // --- ВидНалоговогоОбязательства + СубконтоДт1 (Перечисление.ВидыПлатежейВГосБюджет) ---
        if (!string.IsNullOrWhiteSpace(doc.ObligationType))
        {
            try
            {
                dynamic enumMgr = conn.Перечисления.ВидыПлатежейВГосБюджет;
                var enumVal = enumMgr.GetType().InvokeMember(
                    doc.ObligationType, System.Reflection.BindingFlags.GetProperty,
                    null, (object)enumMgr, null);
                // Set both the document property AND the subconto
                try { newDoc.ВидНалоговогоОбязательства = enumVal; } catch { }
                try
                {
                    newDoc.СубконтоДт1 = enumVal;
                    _logger.LogInformation("    СубконтоДт1 (ВидыПлатежейВГосБюджет) = {0}", doc.ObligationType);
                }
                catch (Exception ex2)
                {
                    _logger.LogWarning("    СубконтоДт1 failed: {0}", ex2.Message);
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning("    ВидыПлатежейВГосБюджет.{0} failed: {1}", doc.ObligationType, ex.Message);
            }
        }

        // --- СчетУчетаРасчетовСКонтрагентом (счёт дебета: 4520, 69.20.1 etc.) ---
        if (!string.IsNullOrWhiteSpace(doc.DebitAccount))
        {
            try
            {
                dynamic chartAcct = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду(doc.DebitAccount);
                if (!chartAcct.Пустая())
                {
                    newDoc.СчетУчетаРасчетовСКонтрагентом = chartAcct;
                    _logger.LogInformation("    СчетУчетаРасчетовСКонтрагентом = {0}", doc.DebitAccount);
                }
                else
                    _logger.LogWarning("    Chart account not found: '{0}'", doc.DebitAccount);
            }
            catch (Exception ex)
            {
                _logger.LogWarning("    Failed to set СчетУчетаРасчетовСКонтрагентом '{0}': {1}", doc.DebitAccount, ex.Message);
            }
        }

        // --- ПодразделениеОрганизации — first non-deleted for this organization ---
        try
        {
            dynamic deptSel = conn.Справочники.ПодразделенияОрганизаций.Выбрать();
            while (deptSel.Следующий())
            {
                if ((bool)deptSel.ПометкаУдаления) continue;
                try
                {
                    // Check if department belongs to our organization
                    dynamic owner = deptSel.Владелец;
                    string ownerName = (string)(owner?.Наименование ?? "");
                    string orgName = (string)(org.Наименование ?? "");
                    if (ownerName == orgName)
                    {
                        newDoc.ПодразделениеОрганизации = deptSel.Ссылка;
                        string deptName = (string)(deptSel.Наименование ?? "");
                        _logger.LogInformation("    ПодразделениеОрганизации = '{0}'", deptName);
                        break;
                    }
                }
                catch { }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning("    Failed to set ПодразделениеОрганизации: {Err}", ex.Message);
        }

        // --- НазначениеПлатежа ---
        try { newDoc.НазначениеПлатежа = doc.Purpose; } catch { }

        // --- Salary: link ПлатежнаяВедомость (created earlier in LoadFile) ---
        if (isDebit && doc.OperationType == "Перечисление заработной платы работнику"
            && _currentPayrollRef != null)
        {
            try
            {
                newDoc.ПлатежнаяВедомость = _currentPayrollRef;
                _logger.LogInformation("    ПлатежнаяВедомость linked");
            }
            catch (Exception ex)
            {
                _logger.LogWarning("    Failed to link ведомость: {Err}", ex.Message);
            }
        }

        // НомерВходящегоДокумента / ДатаВходящегоДокумента — optional, set later if needed

        // Write without posting (Записать, не Провести)
        try
        {
            newDoc.Записать();
        }
        catch (Exception writeEx)
        {
            _logger.LogError("    WRITE FAILED: {Err}\n{Stack}", writeEx.Message, writeEx.ToString());
            throw;
        }

        _logger.LogInformation("  Created {Type} #{Number} {Date} {Amount:F2} {Op} {Counterparty}",
            isDebit ? "Debit" : "Credit",
            doc.Number, doc.Date, doc.Amount,
            doc.OperationType,
            counterpartyName);
        return true;
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
                            if (selection.ПометкаУдаления) continue;

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

    private dynamic? FindBankAccountByNumber(dynamic conn, string account)
    {
        string normalized = NormalizeAccount(account);
        dynamic sel = conn.Справочники.БанковскиеСчета.Выбрать();
        while (sel.Следующий())
        {
            try
            {
                string no = NormalizeAccount(sel.НомерСчета ?? "");
                if (no == normalized)
                    return sel.Ссылка;
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

    // ─── Pre-load: ensure counterparties & bank accounts exist ───
    private static readonly string[] _companyIndicators = {
        "DOO", "D.O.O.", "D.O.O", "AD", "A.D.", "A.D",
        "LLC", "LTD", "GMBH", "SRL", "OOO", "ООО", "ДОО"
    };

    private bool IsCompanyName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return false;
        var upper = name.ToUpperInvariant();
        foreach (var ind in _companyIndicators)
            if (upper.Contains(ind)) return true;
        return false;
    }

    /// <summary>Strip company suffixes (DOO, D.O.O., AD, A.D., etc.) for fuzzy search.</summary>
    private static string StripCompanySuffix(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "";
        // Remove known suffixes (case-insensitive), then trim
        var result = name.Trim();
        foreach (var suffix in new[] { "D.O.O.", "D.O.O", "DOO", "A.D.", "A.D", "AD",
            "LLC", "LTD", "GMBH", "SRL", "OOO", "ООО", "ДОО" })
        {
            // Check at end or as a separate word
            if (result.EndsWith(" " + suffix, StringComparison.OrdinalIgnoreCase))
            {
                result = result[..^(suffix.Length + 1)].Trim();
            }
            else if (result.Equals(suffix, StringComparison.OrdinalIgnoreCase))
            {
                return result; // entire name is a suffix — don't strip
            }
        }
        return result.Trim();
    }

    /// <summary>
    /// Before loading documents, ensure all counterparties and their bank accounts exist in 1C.
    /// 1) Find bank account → done
    /// 2) Bank account not found → find counterparty by name → create bank account
    /// 3) Counterparty not found → create counterparty (company/person) + bank account
    /// </summary>
    private void EnsureCounterparties(dynamic conn, BankExchangeFile file)
    {
        var seen = new HashSet<string>();
        var counterparties = new List<(string Account, string Name, string Inn)>();

        foreach (var doc in file.Documents)
        {
            bool isDebit = doc.IsDebit(file.Account);
            string cpAcct = isDebit ? doc.RecipientAccount : doc.PayerAccount;
            string cpName = isDebit ? doc.RecipientName : doc.PayerName;
            string cpInn = isDebit ? doc.RecipientInn : doc.PayerInn;

            if (string.IsNullOrWhiteSpace(cpAcct)) continue;
            string normAcct = NormalizeAccount(cpAcct);
            if (string.IsNullOrWhiteSpace(normAcct) || normAcct.StartsWith("907")) continue;
            if (!seen.Add(normAcct)) continue;

            counterparties.Add((cpAcct, cpName, cpInn));
        }

        if (counterparties.Count == 0) return;
        _logger.LogInformation("  EnsureCounterparties: checking {Count} unique counterparties", counterparties.Count);

        foreach (var (acct, name, inn) in counterparties)
        {
            try
            {
                string normAcct = NormalizeAccount(acct);

                // Step 1: find bank account
                dynamic? bankAcctRef = FindBankAccountByNumber(conn, acct);
                if (bankAcctRef != null)
                {
                    _logger.LogInformation("    {Acct} — bank account exists", normAcct);
                    continue;
                }

                _logger.LogInformation("    {Acct} ({Name}) — bank account NOT found", normAcct, name);

                // Step 2: find counterparty by INN or name
                dynamic? counterpartyRef = null;

                if (!string.IsNullOrWhiteSpace(inn))
                {
                    try
                    {
                        dynamic found = conn.Справочники.Контрагенты
                            .НайтиПоРеквизиту("ИНН", inn);
                        if (!found.Пустая())
                            counterpartyRef = found;
                    }
                    catch { }
                }

                if (counterpartyRef == null && !string.IsNullOrWhiteSpace(name))
                {
                    try
                    {
                        dynamic found = conn.Справочники.Контрагенты
                            .НайтиПоНаименованию(name, true);
                        if (!found.Пустая())
                            counterpartyRef = found;
                    }
                    catch { }
                }

                // Fallback: for companies, search by name without DOO/AD suffix
                if (counterpartyRef == null && IsCompanyName(name))
                {
                    string stripped = StripCompanySuffix(name);
                    if (!string.IsNullOrWhiteSpace(stripped) && stripped != name)
                    {
                        _logger.LogInformation("    Trying stripped name: '{Stripped}'", stripped);
                        try
                        {
                            // НайтиПоНаименованию with exact=false does prefix search
                            dynamic sel = conn.Справочники.Контрагенты.Выбрать();
                            while (sel.Следующий())
                            {
                                try
                                {
                                    if ((bool)sel.ПометкаУдаления) continue;
                                    string cpName1c = ((string)(sel.Наименование ?? "")).Trim();
                                    string cpStripped = StripCompanySuffix(cpName1c);
                                    if (cpStripped.Equals(stripped, StringComparison.OrdinalIgnoreCase))
                                    {
                                        counterpartyRef = sel.Ссылка;
                                        _logger.LogInformation("    Matched by stripped name: '{Name1C}'", cpName1c);
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }

                // Step 3: create counterparty if not found
                if (counterpartyRef == null)
                {
                    bool isCompany = IsCompanyName(name);
                    _logger.LogInformation("    Creating {Type}: {Name}",
                        isCompany ? "company" : "physical person", name);

                    if (isCompany)
                    {
                        dynamic newCp = conn.Справочники.Контрагенты.СоздатьЭлемент();
                        newCp.Наименование = name;
                        if (!string.IsNullOrWhiteSpace(inn))
                        {
                            try { newCp.ИНН = inn; } catch { }
                        }
                        try
                        {
                            dynamic enumMgr = conn.Перечисления.ЮрФизЛицо;
                            var val = enumMgr.GetType().InvokeMember("ЮрЛицо",
                                System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                            newCp.ЮрФизЛицо = val;
                        }
                        catch { }
                        newCp.Записать();
                        counterpartyRef = newCp.Ссылка;
                        _logger.LogInformation("    Created company: {Name}", name);
                    }
                    else
                    {
                        // Physical person: create ФизическоеЛицо + Контрагент
                        dynamic newPhys = conn.Справочники.ФизическиеЛица.СоздатьЭлемент();
                        newPhys.Наименование = name;
                        newPhys.Записать();
                        _logger.LogInformation("    Created ФизическоеЛицо: {Name}", name);

                        dynamic newCp = conn.Справочники.Контрагенты.СоздатьЭлемент();
                        newCp.Наименование = name;
                        try
                        {
                            dynamic enumMgr = conn.Перечисления.ЮрФизЛицо;
                            var val = enumMgr.GetType().InvokeMember("ФизЛицо",
                                System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                            newCp.ЮрФизЛицо = val;
                        }
                        catch { }
                        newCp.Записать();
                        counterpartyRef = newCp.Ссылка;
                        _logger.LogInformation("    Created person: {Name}", name);
                    }
                }
                else
                {
                    string cpNameFound = "";
                    try { cpNameFound = (string)(counterpartyRef.Наименование ?? ""); } catch { }
                    _logger.LogInformation("    Found counterparty: {Name}", cpNameFound);
                }

                // Step 4: create bank account for the counterparty
                if (counterpartyRef != null)
                {
                    try
                    {
                        string bankCode = normAcct.Length >= 3 ? normAcct.Substring(0, 3) : "";
                        dynamic? bankRef = null;
                        if (!string.IsNullOrWhiteSpace(bankCode))
                        {
                            try
                            {
                                dynamic bankSel = conn.Справочники.Банки.Выбрать();
                                while (bankSel.Следующий())
                                {
                                    try
                                    {
                                        string code = (string)(bankSel.Код ?? "");
                                        if (code == bankCode)
                                        {
                                            bankRef = bankSel.Ссылка;
                                            break;
                                        }
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                        }

                        dynamic newBankAcct = conn.Справочники.БанковскиеСчета.СоздатьЭлемент();
                        newBankAcct.НомерСчета = normAcct;
                        newBankAcct.Владелец = counterpartyRef;
                        if (bankRef != null)
                        {
                            try { newBankAcct.Банк = bankRef; } catch { }
                        }
                        try
                        {
                            dynamic eur = conn.Справочники.Валюты.НайтиПоКоду("978");
                            if (eur.Пустая())
                                eur = conn.Справочники.Валюты.НайтиПоНаименованию("EUR", true);
                            if (!eur.Пустая())
                                newBankAcct.ВалютаДенежныхСредств = eur;
                        }
                        catch { }

                        newBankAcct.Записать();
                        _logger.LogInformation("    Created bank account {Acct} for {Name}", normAcct, name);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning("    Failed to create bank account {Acct}: {Err}", normAcct, ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning("    EnsureCounterparties error for {Acct}: {Err}", acct, ex.Message);
            }
        }
    }

    public void Dispose() { }
}

public class LoadResult
{
    public bool Success { get; set; }
    public int Created { get; set; }
    public int Skipped { get; set; }
    public string? Error { get; set; }
}

// ──────────────────────────────────────────────────────────────
// Background worker service
// ──────────────────────────────────────────────────────────────
public class LoaderWorker : BackgroundService
{
    private readonly LoaderConfig _config;
    private readonly Com1CConnector _com;
    private readonly ILogger<LoaderWorker> _logger;
    private readonly HashSet<string> _processedFiles = new();

    public LoaderWorker(LoaderConfig config, Com1CConnector com,
        ILogger<LoaderWorker> logger)
    {
        _config = config;
        _com = com;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("Loader1C service starting...");

        // Always scan databases on startup to refresh mapping
        await Task.Run(() => _com.ScanDatabases(), stoppingToken);

        // Track already processed files
        var loadedDir = _config.LoadedDir;
        Directory.CreateDirectory(loadedDir);

        // Scan loaded dir to know what's already done
        if (Directory.Exists(loadedDir))
        {
            foreach (var f in Directory.GetFiles(loadedDir, "*.txt",
                         SearchOption.AllDirectories))
                _processedFiles.Add(Path.GetFileName(f));
        }

        // Also check for .loaded marker files in output dir
        if (Directory.Exists(_config.OutputDir))
        {
            foreach (var f in Directory.GetFiles(_config.OutputDir, "*.txt.loaded",
                         SearchOption.AllDirectories))
            {
                // marker "xxx.txt.loaded" means "xxx.txt" was processed
                var originalName = Path.GetFileNameWithoutExtension(f); // strips .loaded → xxx.txt
                _processedFiles.Add(originalName);
            }
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

    private void ScanAndLoad()
    {
        if (!Directory.Exists(_config.OutputDir)) return;

        var txtFiles = Directory.GetFiles(_config.OutputDir, "*.txt",
            SearchOption.AllDirectories);

        foreach (var filePath in txtFiles)
        {
            var fileName = Path.GetFileName(filePath);
            if (_processedFiles.Contains(fileName)) continue;

            // Skip files that look like garbage (unknown_*)
            if (fileName.StartsWith("unknown_")) continue;

            // Skip if .loaded marker exists (already processed)
            if (File.Exists(filePath + ".loaded")) continue;

            _logger.LogInformation("New file: {File}", filePath);

            try
            {
                var parsed = BankFileParser.Parse(filePath);
                if (parsed.Documents.Count == 0)
                {
                    _logger.LogWarning("No documents in {File}, skipping", fileName);
                    _processedFiles.Add(fileName);
                    continue;
                }

                var result = _com.LoadFile(parsed);

                if (result.Created > 0)
                {
                    _logger.LogInformation("Loaded {File}: {Created} documents created",
                        fileName, result.Created);

                    // Write .loaded marker FIRST (survives crash/restart)
                    File.WriteAllText(filePath + ".loaded",
                        $"Loaded: {DateTime.Now:yyyy-MM-dd HH:mm:ss}, " +
                        $"Documents: {result.Created}");

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
                else
                {
                    _logger.LogWarning("Failed to load {File}: {Error}",
                        fileName, result.Error);
                }

                _processedFiles.Add(fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing {File}", filePath);
                _processedFiles.Add(fileName); // don't retry endlessly
            }
        }
    }
}

// ──────────────────────────────────────────────────────────────
// Entry point
// ──────────────────────────────────────────────────────────────
public class Program
{
    public static void Main(string[] args)
    {
        // Register windows-1251 encoding
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

        var config = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: false)
            .Build();

        var loaderConfig = config.GetSection("Loader").Get<LoaderConfig>()
            ?? throw new Exception("Missing 'Loader' section in appsettings.json");

        // CLI commands
        var command = args.Length > 0 ? args[0].ToLower() : "run";

        if (command == "diag")
        {
            // Diagnostic: discover property names for bank documents
            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var logger = loggerFactory.CreateLogger<Program>();

            string db = args.Length > 1 ? args[1] : "base_rs";
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine($"Connecting to {db}...");

            var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                ?? throw new Exception("V83.COMConnector not registered");
            dynamic connector = Activator.CreateInstance(connectorType)!;
            var connStr = $"Srvr=\"{loaderConfig.Server}\";Ref=\"{db}\";" +
                          $"Usr=\"{loaderConfig.User}\";Pwd=\"{loaderConfig.Password}\"";
            dynamic conn = connector.Connect(connStr);
            Console.WriteLine("Connected!");

            // Create a test debit document to inspect its metadata
            Console.WriteLine("\n=== СписаниеСРасчетногоСчета ===");
            dynamic debitDoc = conn.Документы.СписаниеСРасчетногоСчета.СоздатьДокумент();
            dynamic debitMeta = debitDoc.Метаданные();

            Console.WriteLine("Реквизиты:");
            dynamic debitAttrs = debitMeta.Реквизиты;
            for (int i = 0; i < (int)debitAttrs.Количество(); i++)
            {
                dynamic attr = debitAttrs.Получить(i);
                Console.WriteLine($"  {attr.Имя}  ({attr.Тип})");
            }

            Console.WriteLine("\nТабличные части:");
            dynamic debitTables = debitMeta.ТабличныеЧасти;
            for (int i = 0; i < (int)debitTables.Количество(); i++)
            {
                dynamic table = debitTables.Получить(i);
                Console.WriteLine($"  {table.Имя}:");
                dynamic tableCols = table.Реквизиты;
                for (int j = 0; j < (int)tableCols.Количество(); j++)
                {
                    dynamic col = tableCols.Получить(j);
                    Console.WriteLine($"    {col.Имя}  ({col.Тип})");
                }
            }

            // List operation type enums
            Console.WriteLine("\n=== Перечисления (операции) ===");
            dynamic allEnums = conn.Метаданные.Перечисления;
            for (int i = 0; i < (int)allEnums.Количество(); i++)
            {
                dynamic en = allEnums.Получить(i);
                string name = en.Имя;
                if (name.Contains("Операци") || name.Contains("Списани") ||
                    name.Contains("Поступлени") || name.Contains("Банк") ||
                    name.Contains("Налог") || name.Contains("Обязательств") ||
                    name.Contains("Бюджет"))
                {
                    Console.WriteLine($"  {name}:");
                    dynamic vals = en.ЗначенияПеречисления;
                    for (int j = 0; j < (int)vals.Количество(); j++)
                    {
                        Console.WriteLine($"    {vals.Получить(j).Имя}");
                    }
                }
            }

            // List relevant catalogs
            Console.WriteLine("\n=== Справочники (движ.ден.средств, расходы) ===");
            dynamic allCats = conn.Метаданные.Справочники;
            for (int i = 0; i < (int)allCats.Количество(); i++)
            {
                dynamic cat = allCats.Получить(i);
                string catName = cat.Имя;
                if (catName.Contains("Стать") || catName.Contains("Движени") ||
                    catName.Contains("Расход") || catName.Contains("Налог") ||
                    catName.Contains("Подразделен"))
                {
                    Console.WriteLine($"  {catName}");
                }
            }

            // Dump specific catalog values
            // Check the actual type of ВидНалоговогоОбязательства attribute
            Console.WriteLine("\n=== Type of ВидНалоговогоОбязательства ===");
            try
            {
                dynamic vidNalOblAttr = debitMeta.Реквизиты.Найти("ВидНалоговогоОбязательства");
                if (vidNalOblAttr != null)
                {
                    dynamic typeDesc = vidNalOblAttr.Тип;
                    Console.WriteLine($"  TypeDescription: {typeDesc}");
                    dynamic typeArr = typeDesc.Типы();
                    int cnt = (int)typeArr.Количество();
                    Console.WriteLine($"  Types count: {cnt}");
                    for (int i = 0; i < cnt; i++)
                    {
                        dynamic t = typeArr.Получить(i);
                        Console.WriteLine($"  Type[{i}]: {t}");
                        // Try to get type name via metadata
                        try
                        {
                            dynamic md = conn.Метаданные.НайтиПоТипу(t);
                            Console.WriteLine($"    Metadata: {md.ПолноеИмя()}");
                        }
                        catch (Exception ex2) { Console.WriteLine($"    Metadata: {ex2.Message}"); }
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  Error: {ex.Message}"); }

            Console.WriteLine("\n=== Catalog values: ВидыОбязательствНалогоплательщика ===");
            try
            {
                dynamic oblSel2 = conn.Справочники.ВидыОбязательствНалогоплательщика.Выбрать();
                while (oblSel2.Следующий())
                    Console.WriteLine($"  '{oblSel2.Наименование}' (deleted={oblSel2.ПометкаУдаления})");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  (catalog error: {ex.Message})");
                // Try as enum
                try
                {
                    dynamic allEnums3 = conn.Метаданные.Перечисления;
                    for (int i = 0; i < (int)allEnums3.Количество(); i++)
                    {
                        dynamic en = allEnums3.Получить(i);
                        if (en.Имя.ToString().Contains("Обязательств"))
                        {
                            Console.WriteLine($"  Enum: {en.Имя}");
                            dynamic vals = en.ЗначенияПеречисления;
                            for (int j = 0; j < (int)vals.Количество(); j++)
                                Console.WriteLine($"    {vals.Получить(j).Имя}");
                        }
                    }
                }
                catch { }
            }

            Console.WriteLine("\n=== Catalog values: ВидыНалоговыхОбязательств ===");
            try
            {
                dynamic oblSel = conn.Справочники.ВидыНалоговыхОбязательств.Выбрать();
                while (oblSel.Следующий())
                    Console.WriteLine($"  '{oblSel.Наименование}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  (catalog not found, trying enum: {ex.Message})");
                try
                {
                    dynamic allEnums2 = conn.Метаданные.Перечисления;
                    for (int i = 0; i < (int)allEnums2.Количество(); i++)
                    {
                        dynamic en = allEnums2.Получить(i);
                        if (en.Имя.ToString().Contains("Обязательств"))
                        {
                            Console.WriteLine($"  Enum: {en.Имя}");
                            dynamic vals = en.ЗначенияПеречисления;
                            for (int j = 0; j < (int)vals.Количество(); j++)
                                Console.WriteLine($"    {vals.Получить(j).Имя}");
                        }
                    }
                }
                catch { }
            }

            Console.WriteLine("\n=== Catalog values: ПодразделенияОрганизаций ===");
            try
            {
                dynamic deptSel = conn.Справочники.ПодразделенияОрганизаций.Выбрать();
                while (deptSel.Следующий())
                    Console.WriteLine($"  '{deptSel.Наименование}' (deleted={deptSel.ПометкаУдаления})");
            }
            catch (Exception ex) { Console.WriteLine($"  Error: {ex.Message}"); }

            Console.WriteLine("\n=== Catalog values: ВидыНалоговИПлатежейВБюджет ===");
            try
            {
                dynamic taxSel = conn.Справочники.ВидыНалоговИПлатежейВБюджет.Выбрать();
                while (taxSel.Следующий())
                    Console.WriteLine($"  '{taxSel.Наименование}' (deleted={taxSel.ПометкаУдаления})");
            }
            catch (Exception ex) { Console.WriteLine($"  Error: {ex.Message}"); }

            Console.WriteLine("\n=== ПланыСчетов.Хозрасчетный ===");
            string[] testCodes = { "51", "4520", "45.20", "45", "68", "68.10", "69", "69.20", "69.20.1", "91.02" };
            foreach (var code in testCodes)
            {
                try
                {
                    dynamic acct = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду(code);
                    bool isEmpty = acct.Пустая();
                    string name = isEmpty ? "(empty)" : (string)acct.Наименование;
                    Console.WriteLine($"  {code} → {name} (empty={isEmpty})");
                }
                catch (Exception ex) { Console.WriteLine($"  {code} → Error: {ex.Message}"); }
            }

            // Dump ВедомостьНаВыплатуЗарплаты metadata
            Console.WriteLine("\n=== ВедомостьНаВыплатуЗарплаты metadata ===");
            try
            {
                dynamic vedDoc = conn.Документы.ВедомостьНаВыплатуЗарплаты.СоздатьДокумент();
                dynamic vedMeta = vedDoc.Метаданные();
                dynamic vedAttrs = vedMeta.Реквизиты;
                for (int i = 0; i < (int)vedAttrs.Количество(); i++)
                {
                    dynamic attr = vedAttrs.Получить(i);
                    Console.WriteLine($"  {attr.Имя}  ({attr.Тип})");
                    try
                    {
                        dynamic typeArr = attr.Тип.Типы();
                        for (int j = 0; j < (int)typeArr.Количество(); j++)
                        {
                            dynamic t = typeArr.Получить(j);
                            try
                            {
                                dynamic md = conn.Метаданные.НайтиПоТипу(t);
                                string fullName = (string)md.ПолноеИмя();
                                Console.WriteLine($"    -> {fullName}");
                                if (fullName.StartsWith("Перечисление."))
                                {
                                    string enumName = fullName.Replace("Перечисление.", "");
                                    dynamic enumMeta = conn.Метаданные.Перечисления.Найти(enumName);
                                    if (enumMeta != null)
                                    {
                                        dynamic vals = enumMeta.ЗначенияПеречисления;
                                        for (int k = 0; k < (int)vals.Количество(); k++)
                                            Console.WriteLine($"      {vals.Получить(k).Имя}");
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Console.WriteLine($"  Error: {ex.Message}"); }

            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return;
        }

        if (command == "inspect" && args.Length > 2)
        {
            // Inspect a specific document by number: inspect <db> <docNumber>
            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var logger = loggerFactory.CreateLogger<Program>();
            string db = args[1];
            string docNum = args[2];
            Console.OutputEncoding = Encoding.UTF8;
            Console.WriteLine($"Connecting to {db}...");

            var connectorType = Type.GetTypeFromProgID("V83.COMConnector")
                ?? throw new Exception("V83.COMConnector not registered");
            dynamic connector = Activator.CreateInstance(connectorType)!;
            var connStr = $"Srvr=\"{loaderConfig.Server}\";Ref=\"{db}\";" +
                          $"Usr=\"{loaderConfig.User}\";Pwd=\"{loaderConfig.Password}\"";
            dynamic conn = connector.Connect(connStr);
            Console.WriteLine("Connected!");

            // Search in both document types — show ALL matches
            bool found = false;
            foreach (var docType in new[] { "СписаниеСРасчетногоСчета", "ПоступлениеНаРасчетныйСчет" })
            {
                try
                {
                    dynamic mgr = conn.Документы.GetType().InvokeMember(
                        docType, System.Reflection.BindingFlags.GetProperty,
                        null, (object)conn.Документы, null);
                    dynamic sel = mgr.Выбрать();
                    while (sel.Следующий())
                    {
                        string num = (string)(sel.Номер ?? "");
                        if (num.Trim() != docNum.Trim()) continue;
                        found = true;

                        // Get object for full attribute access
                        dynamic docObj = sel.Ссылка.ПолучитьОбъект();
                        Console.WriteLine($"\n=== {docType} #{num} от {sel.Дата} ===");

                        // Dump ALL attributes via metadata
                        dynamic meta = docObj.Метаданные();
                        dynamic attrs = meta.Реквизиты;
                        for (int i = 0; i < (int)attrs.Количество(); i++)
                        {
                            dynamic attr = attrs.Получить(i);
                            string attrName = (string)attr.Имя;
                            try
                            {
                                object val = docObj.GetType().InvokeMember(
                                    attrName, System.Reflection.BindingFlags.GetProperty,
                                    null, (object)docObj, null);
                                string valStr = val?.ToString() ?? "(null)";
                                // For COM objects try to get Наименование or НомерСчета
                                if (valStr == "System.__ComObject")
                                {
                                    try { valStr = ((dynamic)val).Наименование ?? valStr; } catch {
                                        try { valStr = ((dynamic)val).НомерСчета ?? valStr; } catch {
                                            try { valStr = ((dynamic)val).Код ?? valStr; } catch {
                                                try { valStr = $"#{((dynamic)val).Номер} от {((dynamic)val).Дата}"; } catch { }
                                            }
                                        }
                                    }
                                    // For ПлатежнаяВедомость — dump its attributes too
                                    if (attrName == "ПлатежнаяВедомость" && valStr != "(null)")
                                    {
                                        Console.WriteLine($"  {attrName} = {valStr}");
                                        try
                                        {
                                            dynamic vedObj = ((dynamic)val).ПолучитьОбъект();
                                            dynamic vedMeta = vedObj.Метаданные();
                                            Console.WriteLine($"    >>> Тип: {vedMeta.ПолноеИмя()}");
                                            dynamic vedAttrs = vedMeta.Реквизиты;
                                            for (int va = 0; va < (int)vedAttrs.Количество(); va++)
                                            {
                                                dynamic vattr = vedAttrs.Получить(va);
                                                string vName = (string)vattr.Имя;
                                                try
                                                {
                                                    object vv = vedObj.GetType().InvokeMember(
                                                        vName, System.Reflection.BindingFlags.GetProperty,
                                                        null, (object)vedObj, null);
                                                    string vvStr = vv?.ToString() ?? "(null)";
                                                    if (vvStr == "System.__ComObject")
                                                        try { vvStr = ((dynamic)vv).Наименование ?? vvStr; } catch {
                                                            try { vvStr = $"#{((dynamic)vv).Номер} от {((dynamic)vv).Дата}"; } catch { }
                                                        }
                                                    if (vvStr.Length > 100) vvStr = vvStr[..100] + "...";
                                                    Console.WriteLine($"    {vName} = {vvStr}");
                                                }
                                                catch { }
                                            }
                                            // Tabular sections of the vedомость
                                            dynamic vedTables = vedMeta.ТабличныеЧасти;
                                            for (int vt = 0; vt < (int)vedTables.Количество(); vt++)
                                            {
                                                dynamic vtable = vedTables.Получить(vt);
                                                string vtName = (string)vtable.Имя;
                                                try
                                                {
                                                    dynamic vtSection = vedObj.GetType().InvokeMember(
                                                        vtName, System.Reflection.BindingFlags.GetProperty,
                                                        null, (object)vedObj, null);
                                                    int vtRows = (int)vtSection.Количество();
                                                    Console.WriteLine($"    --- {vtName}: {vtRows} rows ---");
                                                    dynamic vtCols = vtable.Реквизиты;
                                                    int vtColCount = (int)vtCols.Количество();
                                                    for (int vr = 0; vr < Math.Min(vtRows, 10); vr++)
                                                    {
                                                        dynamic vrow = vtSection.Получить(vr);
                                                        Console.Write($"      row[{vr}]: ");
                                                        for (int vc = 0; vc < vtColCount; vc++)
                                                        {
                                                            string vcName = (string)vtCols.Получить(vc).Имя;
                                                            try
                                                            {
                                                                object vcv = vrow.GetType().InvokeMember(
                                                                    vcName, System.Reflection.BindingFlags.GetProperty,
                                                                    null, (object)vrow, null);
                                                                string vcvStr = vcv?.ToString() ?? "";
                                                                if (vcvStr == "System.__ComObject")
                                                                    try { vcvStr = ((dynamic)vcv).Наименование ?? vcvStr; } catch { }
                                                                Console.Write($"{vcName}={vcvStr}; ");
                                                            }
                                                            catch { }
                                                        }
                                                        Console.WriteLine();
                                                    }
                                                }
                                                catch { }
                                            }
                                        }
                                        catch (Exception vex) { Console.WriteLine($"    >>> Error: {vex.Message}"); }
                                        continue; // already printed
                                    }
                                }
                                if (valStr.Length > 120) valStr = valStr[..120] + "...";
                                Console.WriteLine($"  {attrName} = {valStr}");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"  {attrName} = ERROR: {ex.Message}");
                            }
                        }

                        // Tabular sections
                        dynamic tables = meta.ТабличныеЧасти;
                        for (int ti = 0; ti < (int)tables.Количество(); ti++)
                        {
                            dynamic table = tables.Получить(ti);
                            string tableName = (string)table.Имя;
                            try
                            {
                                dynamic tabSection = docObj.GetType().InvokeMember(
                                    tableName, System.Reflection.BindingFlags.GetProperty,
                                    null, (object)docObj, null);
                                int rowCount = (int)tabSection.Количество();
                                Console.WriteLine($"\n  --- {tableName}: {rowCount} rows ---");
                                if (rowCount > 0)
                                {
                                    dynamic tableCols = table.Реквизиты;
                                    int colCount = (int)tableCols.Количество();
                                    for (int r = 0; r < Math.Min(rowCount, 5); r++)
                                    {
                                        dynamic row = tabSection.Получить(r);
                                        Console.Write($"    row[{r}]: ");
                                        for (int c = 0; c < colCount; c++)
                                        {
                                            string colName = (string)tableCols.Получить(c).Имя;
                                            try
                                            {
                                                object cv = row.GetType().InvokeMember(
                                                    colName, System.Reflection.BindingFlags.GetProperty,
                                                    null, (object)row, null);
                                                string cvStr = cv?.ToString() ?? "";
                                                if (cvStr == "System.__ComObject")
                                                    try { cvStr = ((dynamic)cv).Наименование ?? cvStr; } catch { }
                                                Console.Write($"{colName}={cvStr}; ");
                                            }
                                            catch { Console.Write($"{colName}=?; "); }
                                        }
                                        Console.WriteLine();
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                }
                catch { }
            }
            if (!found) Console.WriteLine($"Document {docNum} not found in {db}");
            Marshal.ReleaseComObject(conn);
            Marshal.ReleaseComObject(connector);
            return;
        }

        if (command == "scan")
        {
            // Just scan databases and save mapping
            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var logger = loggerFactory.CreateLogger<Program>();
            var com = new Com1CConnector(loaderConfig,
                loggerFactory.CreateLogger("COM"));
            com.ScanDatabases();
            Console.WriteLine($"Mapping saved. {com.AccountMap.Count} accounts found.");
            return;
        }

        if (command == "test" && args.Length > 1)
        {
            // Test parse a file (dry run, no loading)
            var filePath = args[1];
            var parsed = BankFileParser.Parse(filePath);
            Console.WriteLine($"File: {filePath}");
            Console.WriteLine($"Account: {parsed.Account}");
            Console.WriteLine($"Period: {parsed.DateStart} - {parsed.DateEnd}");
            Console.WriteLine($"Documents: {parsed.Documents.Count}");
            foreach (var doc in parsed.Documents)
            {
                var dir = doc.IsDebit(parsed.Account) ? "DEBIT" : "CREDIT";
                Console.WriteLine($"  {dir} #{doc.Number} {doc.Date} " +
                    $"{doc.Amount:F2} {doc.Purpose[..Math.Min(60, doc.Purpose.Length)]}");
            }

            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var com = new Com1CConnector(loaderConfig,
                loggerFactory.CreateLogger("COM"));
            com.LoadCachedMapping();
            var mapping = com.FindAccount(parsed.Account);
            if (mapping != null)
                Console.WriteLine($"\nMapped to: {mapping.Database} ({mapping.OrgName})");
            else
                Console.WriteLine($"\nWARNING: Account {parsed.Account} not found in mapping!");
            return;
        }

        if (command == "load" && args.Length > 1)
        {
            // Load a single file
            var filePath = args[1];
            using var loggerFactory = LoggerFactory.Create(b => b.AddConsole());
            var logger = loggerFactory.CreateLogger("COM");
            var com = new Com1CConnector(loaderConfig, logger);
            if (!com.LoadCachedMapping())
            {
                Console.WriteLine("No cached mapping. Run 'scan' first.");
                return;
            }
            var parsed = BankFileParser.Parse(filePath);
            var result = com.LoadFile(parsed);
            Console.WriteLine(result.Success
                ? $"OK: {result.Created} documents created"
                : $"FAILED: {result.Error}");
            return;
        }

        // Default: run as service
        var builder = Host.CreateApplicationBuilder(args);
        builder.Services.AddSingleton(loaderConfig);
        builder.Services.AddSingleton(sp =>
            new Com1CConnector(loaderConfig,
                sp.GetRequiredService<ILogger<Com1CConnector>>()));
        builder.Services.AddHostedService<LoaderWorker>();
        builder.Services.AddWindowsService(options =>
            options.ServiceName = "Loader1C");

        // Configure logging: console + file
        builder.Logging.ClearProviders();
        builder.Logging.AddConsole();
        if (!string.IsNullOrEmpty(loaderConfig.LogFile))
        {
            var logDir = Path.GetDirectoryName(loaderConfig.LogFile);
            if (!string.IsNullOrEmpty(logDir))
                Directory.CreateDirectory(logDir);
            builder.Logging.AddProvider(new FileLoggerProvider(loaderConfig.LogFile));
        }

        var host = builder.Build();
        host.Run();
    }
}

// ──────────────────────────────────────────────────────────────
// Simple file logger (no external dependencies)
// ──────────────────────────────────────────────────────────────
public class FileLoggerProvider : ILoggerProvider
{
    private readonly string _filePath;
    private readonly object _lock = new();

    public FileLoggerProvider(string filePath) => _filePath = filePath;

    public ILogger CreateLogger(string categoryName) =>
        new FileLogger(_filePath, categoryName, _lock);

    public void Dispose() { }
}

public class FileLogger : ILogger
{
    private readonly string _filePath;
    private readonly string _category;
    private readonly object _lock;

    public FileLogger(string filePath, string category, object lockObj)
    {
        _filePath = filePath;
        _category = category;
        _lock = lockObj;
    }

    public IDisposable? BeginScope<TState>(TState state) where TState : notnull => null;
    public bool IsEnabled(LogLevel logLevel) => logLevel >= LogLevel.Information;

    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
        Exception? exception, Func<TState, Exception?, string> formatter)
    {
        if (!IsEnabled(logLevel)) return;
        var msg = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{logLevel}] {_category}: {formatter(state, exception)}";
        if (exception != null)
            msg += Environment.NewLine + exception;
        lock (_lock)
        {
            File.AppendAllText(_filePath, msg + Environment.NewLine, Encoding.UTF8);
        }
    }
}
