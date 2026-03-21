using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

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
