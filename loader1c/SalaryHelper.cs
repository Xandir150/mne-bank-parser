using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.Extensions.Logging;

// ──────────────────────────────────────────────────────────────
// Salary-related helpers (partial class of Com1CConnector)
// ──────────────────────────────────────────────────────────────
public partial class Com1CConnector
{
    /// <summary>
    /// Create a bank account for a ФизическоеЛицо (or any owner).
    /// Sets: НомерСчета, Владелец, Банк (by first 3 digits), Валюта, ВидСчета.
    /// </summary>
    public void CreateBankAccountForOwner(dynamic conn, string account, dynamic ownerRef, string ownerName)
    {
        string normAcct = NormalizeAccount(account);
        string bankCode = normAcct.Length >= 3 ? normAcct.Substring(0, 3) : "";

        dynamic newBankAcct = conn.Справочники.БанковскиеСчета.СоздатьЭлемент();
        newBankAcct.НомерСчета = normAcct;
        newBankAcct.Владелец = ownerRef;

        // Bank by code (first 3 digits)
        if (!string.IsNullOrWhiteSpace(bankCode))
        {
            try
            {
                dynamic bankSel = conn.Справочники.Банки.Выбрать();
                while (bankSel.Следующий())
                {
                    string code = ((string)(bankSel.Код ?? "")).Trim();
                    if (code == bankCode)
                    {
                        newBankAcct.Банк = bankSel.Ссылка;
                        break;
                    }
                }
            }
            catch { }
        }

        // Currency = EUR
        try
        {
            dynamic eur = conn.Справочники.Валюты.НайтиПоНаименованию("EUR", true);
            if (!eur.Пустая())
                newBankAcct.ВалютаДенежныхСредств = eur;
        }
        catch { }

        // ВидСчета = copy from existing
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
                        break;
                    }
                }
                catch { }
            }
        }
        catch { }

        newBankAcct.Записать();
        _logger.LogInformation("    Created bank account {Acct} for {Name}", normAcct, ownerName);
    }
}
