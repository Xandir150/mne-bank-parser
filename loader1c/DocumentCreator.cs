using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.Extensions.Logging;

// ──────────────────────────────────────────────────────────────
// Document creation logic (partial class of Com1CConnector)
// ──────────────────────────────────────────────────────────────
public partial class Com1CConnector
{
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
        ["Перевод на другой счёт организации"] = ("ПереводНаДругойСчет", "ПереводСДругогоСчета"),
        ["Возврат от поставщика"] = ("ОплатаПоставщику", "ВозвратОтПоставщика"),
    };

    /// <summary>Returns true if document was created.</summary>
    private bool CreateDocument(dynamic conn, BankDocument doc,
        dynamic bankAcct, dynamic org, bool isDebit,
        string ourAccount = "")
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

        // Check for internal transfer: counterparty account belongs to SAME organization (not just same DB)
        bool isInternalTransfer = false;
        {
            string cpAcctCheck = isDebit
                ? NormalizeAccount(doc.RecipientAccount)
                : NormalizeAccount(doc.PayerAccount);
            if (!string.IsNullOrWhiteSpace(ourAccount) && !string.IsNullOrWhiteSpace(cpAcctCheck))
            {
                var ourMap = FindAccount(ourAccount);
                var cpMap = FindAccount(cpAcctCheck);
                string ourNorm = NormalizeAccount(ourAccount);

                if (ourMap != null && cpMap != null
                    && ourMap.Database == cpMap.Database
                    && !string.IsNullOrWhiteSpace(ourMap.OrgName)
                    && ourMap.OrgName == cpMap.OrgName
                    && cpAcctCheck != ourNorm)
                {
                    // Different account, same org → internal transfer
                    isInternalTransfer = true;
                    doc.OperationType = "Перевод на другой счёт организации";
                    doc.CashFlowItem = "Внутреннее перемещение денежных средств";
                    _logger.LogInformation("    Internal transfer detected: {Our} → {Cp}",
                        ourAccount, cpAcctCheck);
                }
                else if (cpAcctCheck == ourNorm && isDebit
                    && doc.OperationType != "Комиссия банка"
                    && doc.OperationType != "Прочее списание"
                    && doc.OperationType != "Уплата налога")
                {
                    // Same account as counterparty → transfer to IBAN (ME25) version of same account
                    // (international payment via own account)
                    isInternalTransfer = true;
                    doc.OperationType = "Перевод на другой счёт организации";
                    doc.CashFlowItem = "Внутреннее перемещение денежных средств";
                    if (isDebit) doc.RecipientAccount = "ME25" + ourNorm;
                    else doc.PayerAccount = "ME25" + ourNorm;
                    _logger.LogInformation("    International transfer detected: {Our} → ME25{Our}",
                        ourAccount, ourNorm);
                }
            }
        }

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
            dynamic eur = conn.Справочники.Валюты.НайтиПоНаименованию("EUR", true);
            if (!eur.Пустая())
            {
                newDoc.ВалютаДокумента = eur;
                _logger.LogInformation("    ВалютаДокумента = {Name} (code={Code})",
                    (string)(eur.Наименование ?? ""), (string)(eur.Код ?? ""));
            }
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

        // --- Counterparty / Internal transfer ---
        string counterpartyName = isDebit ? doc.RecipientName : doc.PayerName;
        string counterpartyInn = isDebit ? doc.RecipientInn : doc.PayerInn;
        string counterpartyAcct = isDebit ? doc.RecipientAccount : doc.PayerAccount;

        if (isInternalTransfer)
        {
            // Internal transfer: СчетКонтрагента = destination bank account, no Контрагент
            if (!string.IsNullOrWhiteSpace(counterpartyAcct))
            {
                try
                {
                    dynamic cpBankAcct = FindBankAccountByNumber(conn, counterpartyAcct);
                    if (cpBankAcct != null)
                    {
                        try { newDoc.СчетКонтрагента = cpBankAcct;
                              _logger.LogInformation("    Destination bank account (СчетКонтрагента): {Acct}", counterpartyAcct);
                        } catch (Exception ex) {
                            _logger.LogWarning("    Failed to set СчетКонтрагента: {Err}", ex.Message);
                        }
                    }
                    else
                    {
                        _logger.LogWarning("    Destination bank account NOT found: {Acct}", counterpartyAcct);
                    }
                }
                catch (Exception ex) { _logger.LogWarning("    Error finding destination account: {Err}", ex.Message); }
            }

            // СчетУчетаРасчетовСКонтрагентом = 57.01
            try
            {
                dynamic transitAcct = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду("57.01");
                if (!transitAcct.Пустая())
                {
                    string foundCode = (string)(transitAcct.Код ?? "");
                    string foundName = (string)(transitAcct.Наименование ?? "");
                    _logger.LogInformation("    НайтиПоКоду(57.01) → code={Code}, name={Name}", foundCode, foundName);
                    newDoc.СчетУчетаРасчетовСКонтрагентом = transitAcct;
                    _logger.LogInformation("    СчетУчетаРасчетовСКонтрагентом set OK");
                }
            }
            catch (Exception ex) { _logger.LogWarning("    Failed to set transit account: {Err}", ex.Message); }

            // ПодразделениеДт (from cached lookup)
            if (_cachedDeptRef != null)
            {
                try { newDoc.ПодразделениеДт = _cachedDeptRef;
                      _logger.LogInformation("    ПодразделениеДт = '{0}'", _cachedDeptName);
                } catch { }
            }
        }
        else
        {
            bool isSalary = doc.OperationType == "Перечисление заработной платы работнику";
            // Find counterparty bank account first (needed for both СчетКонтрагента and owner lookup)
            dynamic cpBankAcctRef = null;
            if (!string.IsNullOrWhiteSpace(counterpartyAcct))
            {
                try
                {
                    _logger.LogInformation("    Looking for bank account: {Acct}", counterpartyAcct);
                    cpBankAcctRef = FindBankAccountByNumber(conn, counterpartyAcct);
                    if (cpBankAcctRef != null)
                    {
                        _logger.LogInformation("    Found bank account: {Acct}", counterpartyAcct);
                        try { newDoc.СчетКонтрагента = cpBankAcctRef; } catch { }
                    }
                    else
                    {
                        _logger.LogWarning("    Bank account NOT found: {Acct}", counterpartyAcct);
                    }
                }
                catch (Exception ex) { _logger.LogWarning("    Bank account search error: {Err}", ex.Message); }
            }

            // Search counterparty: 1) bank account owner, 2) INN, 3) name with word swap
            bool counterpartySet = false;

            // 1) Bank account owner — most reliable (exact account match)
            // Skip if counterparty account == our account (bank fee charged to our own account)
            string cpNorm = NormalizeAccount(counterpartyAcct);
            string ourNorm2 = NormalizeAccount(ourAccount);
            if (!counterpartySet && cpBankAcctRef != null && cpNorm != ourNorm2)
            {
                try
                {
                    dynamic owner = cpBankAcctRef.Владелец;
                    if (owner != null && !owner.Пустая() && !(bool)owner.ПометкаУдаления)
                    {
                        newDoc.Контрагент = owner;
                        counterpartySet = true;
                        _logger.LogInformation("    Counterparty from bank account: {Name}",
                            (string)(owner.Наименование ?? ""));
                    }
                }
                catch (Exception ex) { _logger.LogWarning("    Bank account owner lookup error: {Err}", ex.Message); }
            }

            // 2) By INN
            if (!counterpartySet && !string.IsNullOrWhiteSpace(counterpartyInn))
            {
                try
                {
                    dynamic found = conn.Справочники.Контрагенты
                        .НайтиПоРеквизиту("ИНН", counterpartyInn);
                    if (!found.Пустая() && !(bool)found.ПометкаУдаления)
                    {
                        newDoc.Контрагент = found;
                        counterpartySet = true;
                        _logger.LogInformation("    Counterparty found by INN: {Inn}", counterpartyInn);
                    }
                }
                catch { }
            }

            // 3) By name (with word-order swap for persons)
            if (!counterpartySet && !string.IsNullOrWhiteSpace(counterpartyName))
            {
                _logger.LogInformation("    Searching counterparty by name: '{Name}'", counterpartyName);
                dynamic? found = FindCounterpartyByName(conn, counterpartyName);
                if (found != null)
                {
                    newDoc.Контрагент = found;
                    counterpartySet = true;
                    _logger.LogInformation("    Counterparty found by name: '{Name}'", counterpartyName);
                }
                else
                {
                    _logger.LogWarning("    Counterparty NOT found by name: '{Name}'", counterpartyName);
                }
            }
            if (!counterpartySet)
            {
                _logger.LogWarning("    Counterparty NOT SET for doc #{Num}", doc.Number);
            }
        }

        // --- ДоговорКонтрагента: find first contract for this counterparty ---
        dynamic? contractRef = null;
        if (!isInternalTransfer && doc.OperationType != "Комиссия банка"
            && doc.OperationType != "Перечисление заработной платы работнику")
        {
            try
            {
                dynamic cpRef = newDoc.Контрагент;
                if (cpRef != null && !cpRef.Пустая())
                {
                    // НайтиПоНаименованию with prefix "BB" for contract named "BB ..."
                    // Or select by owner: Выбрать(owner)
                    try
                    {
                        dynamic contractSel = conn.Справочники.ДоговорыКонтрагентов
                            .GetType().InvokeMember("Выбрать",
                                System.Reflection.BindingFlags.InvokeMethod,
                                null, conn.Справочники.ДоговорыКонтрагентов,
                                new object[] { null, cpRef });
                        while (contractSel.Следующий())
                        {
                            if ((bool)contractSel.ПометкаУдаления) continue;
                            contractRef = contractSel.Ссылка;
                            string cName = (string)(contractSel.Наименование ?? "");
                            newDoc.ДоговорКонтрагента = contractRef;
                            _logger.LogInformation("    ДоговорКонтрагента = '{0}'", cName);
                            break;
                        }
                    }
                    catch
                    {
                        // Fallback: simple select and filter
                        dynamic contractSel = conn.Справочники.ДоговорыКонтрагентов.Выбрать();
                        while (contractSel.Следующий())
                        {
                            try
                            {
                                if ((bool)contractSel.ПометкаУдаления) continue;
                                dynamic owner = contractSel.Владелец;
                                if (owner != null && !owner.Пустая())
                                {
                                    string ownerName = (string)(owner.Наименование ?? "");
                                    string cpName2 = (string)(cpRef.Наименование ?? "");
                                    if (ownerName == cpName2)
                                    {
                                        contractRef = contractSel.Ссылка;
                                        string cName = (string)(contractSel.Наименование ?? "");
                                        newDoc.ДоговорКонтрагента = contractRef;
                                        _logger.LogInformation("    ДоговорКонтрагента = '{0}'", cName);
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }
                    }
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

        // --- СубконтоДт1 for "Прочее списание" (e.g. charity SPASI-ME → Благотворительность) ---
        if (doc.OperationType == "Прочее списание")
        {
            string cpName = isDebit ? (doc.RecipientName ?? "") : (doc.PayerName ?? "");
            if (cpName.IndexOf("spasi-me", StringComparison.OrdinalIgnoreCase) >= 0
                || cpName.IndexOf("spasi me", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                try
                {
                    dynamic item = conn.Справочники.ПрочиеДоходыИРасходы
                        .НайтиПоНаименованию("Благотворительность", false);
                    if (!item.Пустая())
                    {
                        newDoc.СубконтоДт1 = item;
                        _logger.LogInformation("    СубконтоДт1 = Благотворительность");
                    }
                }
                catch { }
            }
        }

        // --- ПодразделениеОрганизации (from cached lookup) ---
        if (_cachedDeptRef != null)
        {
            try { newDoc.ПодразделениеОрганизации = _cachedDeptRef;
                  _logger.LogInformation("    ПодразделениеОрганизации = '{0}'", _cachedDeptName);
            } catch { }
        }

        // --- НазначениеПлатежа ---
        try { newDoc.НазначениеПлатежа = doc.Purpose; } catch { }

        // --- РасшифровкаПлатежа with НДС 21% for supplier payments and returns ---
        if ((isDebit && (doc.OperationType == "Оплата поставщику" || doc.OperationType == ""))
            || doc.OperationType == "Возврат от поставщика")
        {
            try
            {
                dynamic tabPayment = newDoc.РасшифровкаПлатежа;
                dynamic row = tabPayment.Добавить();
                row.СуммаПлатежа = doc.Amount;
                row.КурсВзаиморасчетов = 1;
                row.СуммаВзаиморасчетов = doc.Amount;
                row.КратностьВзаиморасчетов = 1;

                // СпособПогашенияЗадолженности = Автоматически
                try
                {
                    dynamic enumMgr = conn.Перечисления.СпособыПогашенияЗадолженности;
                    var auto = enumMgr.GetType().InvokeMember("Автоматически",
                        System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                    row.СпособПогашенияЗадолженности = auto;
                }
                catch { }

                // НДС 21% (enum НДС20 = переименован в 21% в этой конфигурации)
                try
                {
                    dynamic enumMgr = conn.Перечисления.СтавкиНДС;
                    var nds = enumMgr.GetType().InvokeMember("НДС20",
                        System.Reflection.BindingFlags.GetProperty, null, (object)enumMgr, null);
                    row.СтавкаНДС = nds;
                    // Сумма НДС = amount * 21 / 121
                    row.СуммаНДС = Math.Round(doc.Amount * 21m / 121m, 2);
                }
                catch { }

                // Договор
                if (contractRef != null)
                    try { row.ДоговорКонтрагента = contractRef; } catch { }

                // СтатьяДвиженияДенежныхСредств
                if (!string.IsNullOrWhiteSpace(doc.CashFlowItem))
                {
                    try
                    {
                        dynamic item = conn.Справочники
                            .СтатьиДвиженияДенежныхСредств
                            .НайтиПоНаименованию(doc.CashFlowItem, true);
                        if (!item.Пустая())
                            row.СтатьяДвиженияДенежныхСредств = item;
                    }
                    catch { }
                }

                // СчетУчетаРасчетовСКонтрагентом = 60.01 (Расчеты с поставщиками)
                try
                {
                    dynamic acct60 = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду("60.01");
                    if (!acct60.Пустая())
                        row.СчетУчетаРасчетовСКонтрагентом = acct60;
                }
                catch { }

                // СчетУчетаРасчетовПоАвансам = 60.02
                try
                {
                    dynamic acct6002 = conn.ПланыСчетов.Хозрасчетный.НайтиПоКоду("60.02");
                    if (!acct6002.Пустая())
                        row.СчетУчетаРасчетовПоАвансам = acct6002;
                }
                catch { }
            }
            catch (Exception ex) { _logger.LogWarning("    РасшифровкаПлатежа error: {Err}", ex.Message); }
        }

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

    // ─── Pre-load: ensure counterparties & bank accounts exist ───
    private static readonly string[] _companyIndicators = {
        "DOO", "D.O.O.", "D.O.O", "AD", "A.D.", "A.D",
        "LLC", "LTD", "GMBH", "SRL", "OOO", "ООО", "ДОО",
        "BANK", "BANKA", "BANKE"
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
            "LLC", "LTD", "GMBH", "SRL", "OOO", "ООО", "ДОО",
            "BANKA", "BANK", "BANKE" })
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
            if (string.IsNullOrWhiteSpace(normAcct)) continue;
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
                        if (!found.Пустая() && !(bool)found.ПометкаУдаления)
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
                        if (!found.Пустая() && !(bool)found.ПометкаУдаления)
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
                // Only create new counterparties for payment operations
                if (counterpartyRef == null)
                {
                    bool shouldCreate = file.Documents.Any(d =>
                    {
                        string da = d.IsDebit(file.Account) ? d.RecipientAccount : d.PayerAccount;
                        if (NormalizeAccount(da) != normAcct) return false;
                        string op = d.OperationType ?? "";
                        return op == "Оплата поставщику" || op == "" || op == "Оплата от покупателя";
                    });

                    if (!shouldCreate)
                    {
                        _logger.LogInformation("    Skipping counterparty creation (non-payment op): {Name}", name);
                        // Don't create counterparty, but still can't create bank account without owner
                        continue;
                    }

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
                                        string code = ((string)(bankSel.Код ?? "")).Trim();
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
                            try {
                                newBankAcct.Банк = bankRef;
                                string bankName = "";
                                try { bankName = (string)(bankRef.Наименование ?? ""); } catch { }
                                _logger.LogInformation("    Bank set: {Code} ({Name})", bankCode, bankName);
                            } catch (Exception ex) {
                                _logger.LogWarning("    Failed to set bank: {Err}", ex.Message);
                            }
                        }
                        else
                        {
                            _logger.LogWarning("    Bank not found for code: {Code}", bankCode);
                        }
                        try
                        {
                            dynamic eur = conn.Справочники.Валюты.НайтиПоНаименованию("EUR", true);
                            if (!eur.Пустая())
                                newBankAcct.ВалютаДенежныхСредств = eur;
                        }
                        catch { }

                        // ВидСчета = copy from existing bank account (Расчетный)
                        try
                        {
                            dynamic goodSel = conn.Справочники.БанковскиеСчета.Выбрать();
                            while (goodSel.Следующий())
                            {
                                try
                                {
                                    if ((bool)goodSel.ПометкаУдаления) continue;
                                    dynamic goodObj = goodSel.Ссылка.ПолучитьОбъект();
                                    dynamic goodVid = goodObj.ВидСчета;
                                    string vidStr = goodVid?.ToString() ?? "";
                                    if (!string.IsNullOrWhiteSpace(vidStr) && vidStr != "System.__ComObject")
                                    {
                                        newBankAcct.ВидСчета = goodVid;
                                        _logger.LogInformation("    ВидСчета = {Vid}", vidStr);
                                        break;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch (Exception ex) { _logger.LogWarning("    Failed to set ВидСчета: {Err}", ex.Message); }

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
}
