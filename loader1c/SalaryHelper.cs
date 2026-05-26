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

        // Bank by code (first 3 digits) — cached
        try
        {
            dynamic? bankRef = GetBankByCode(conn, bankCode);
            if (bankRef != null) newBankAcct.Банк = bankRef;
        }
        catch { }

        // Currency = EUR — cached
        try
        {
            dynamic? eur = GetEur(conn);
            if (eur != null) newBankAcct.ВалютаДенежныхСредств = eur;
        }
        catch { }

        // ВидСчета = copy from cached sample
        try
        {
            dynamic? sampleVid = GetSampleVidSchets(conn);
            if (sampleVid != null) newBankAcct.ВидСчета = sampleVid;
        }
        catch { }

        newBankAcct.Записать();
        // Cache the newly created bank account
        try
        {
            dynamic newRef = newBankAcct.Ссылка;
            _cacheBankAcctsByNumber[normAcct] = newRef;
            TrackCom(newRef);
        }
        catch { }
        _logger.LogInformation("    Created bank account {Acct} for {Name}", normAcct, ownerName);
    }
}
