import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


@register_parser
class AdriaticParser(BankParser):
    """Parser for Adriatic Bank (580) PDF statements.

    The bank issues the same layout in two languages and both are in active
    use, so every label is matched in either one:
      English      "STATEMENT TURNOVER" / DATE, TRANSACTION DESCRIPTION,
                   CHARGED, IN BENEFIT
      Montenegrin  "IZVOD PROMETA"      / DATUM, OPIS TRANSAKCIJE,
                   NA TERET, U KORIST
    Only the labels differ — the table column layout is identical.
    Transaction description is multi-line with purpose, reference, counterparty, account.
    """

    bank_code = "580"
    bank_name = "Adriatic Bank"

    def parse(self, file_path: Path) -> ParsedStatement:
        stmt = ParsedStatement(
            bank_code=self.bank_code,
            bank_name=self.bank_name,
        )

        with pdfplumber.open(file_path) as pdf:
            self._parse_header(pdf.pages[0], stmt)
            for page in pdf.pages:
                self._parse_transactions(page, stmt)

        return stmt

    # Label alternations: English | Montenegrin. Serbian diacritics are
    # matched with their ASCII counterpart too — some exports drop them.
    _L_STATEMENT_NO = r"(?:Statement\s+no|Broj\s+izvoda)"
    _L_ACCOUNT_NO = r"(?:Account\s+no|Broj\s+ra[čc]una)"
    _L_CURRENCY = r"(?:Currency|Deviza)"
    _L_STMT_DATE = r"(?:Statem\.\s*date|Datum\s+izvoda)"
    _L_PERIOD = r"(?:For\s+period|Za\s+period)"
    _L_INITIAL = r"(?:INITIAL\s+STATE\s+ON\s+DAY|PO[ČC]ETNO\s+STANJE\s+NA\s+DAN)"

    def _parse_header(self, page, stmt: ParsedStatement) -> None:
        text = page.extract_text() or ""

        # Statement number
        m = re.search(rf"{self._L_STATEMENT_NO}\s*:\s*(\d+)", text)
        if m:
            stmt.statement_number = m.group(1)

        # Account number
        m = re.search(rf"{self._L_ACCOUNT_NO}\s*:\s*(\d+)", text)
        if m:
            stmt.account_number = m.group(1)

        # Currency
        m = re.search(rf"{self._L_CURRENCY}\s*:\s*\d+\s+(\w+)", text)
        if m:
            stmt.currency = m.group(1)

        # Statement date
        m = re.search(rf"{self._L_STMT_DATE}\s*:\s*(\d{{2}}\.\d{{2}}\.\d{{4}})", text)
        if m:
            stmt.statement_date = self.parse_date_dmy(m.group(1))

        # IBAN
        m = re.search(r"IBAN\s*:\s*(ME\d+)", text)
        if m:
            stmt.iban = m.group(1)

        # Client name — trails the date range on the period line, e.g.
        # "For period: 18.02.2026-18.02.2026 BUREVESTNIK MONTENEGRO"
        # "Za period: 31.08.2026-31.08.2026 MINDER DOO"
        m = re.search(
            rf"{self._L_PERIOD}:\s*\d{{2}}\.\d{{2}}\.\d{{4}}\s*-\s*\d{{2}}\.\d{{2}}\.\d{{4}}\s+(.+)",
            text,
        )
        if m:
            stmt.client_name = self.clean_text(m.group(1))

        # Period: "<label>: DD.MM.YYYY-DD.MM.YYYY"
        m = re.search(
            rf"{self._L_PERIOD}:\s*(\d{{2}}\.\d{{2}}\.\d{{4}})\s*-\s*(\d{{2}}\.\d{{2}}\.\d{{4}})",
            text,
        )
        if m:
            stmt.period_start = self.parse_date_dmy(m.group(1))
            stmt.period_end = self.parse_date_dmy(m.group(2))

        # Opening balance: "INITIAL STATE ON DAY: 18.02.2026  203.55"
        #                  "POČETNO STANJE NA DAN: 31.08.2026 2,075.28"
        m = re.search(
            rf"{self._L_INITIAL}:\s*\d{{2}}\.\d{{2}}\.\d{{4}}\s+([\d,]+\.\d{{2}})",
            text,
        )
        if m:
            stmt.opening_balance = self.parse_amount_us(m.group(1))

    def _parse_transactions(self, page, stmt: ParsedStatement) -> None:
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                if not row:
                    continue

                # Skip header rows
                first = (row[0] or "").strip()
                if first in ("DATE", "DATUM", "") and any(
                    ("TRANSACTION" in (c or "") or "OPIS TRANSAKCIJE" in (c or ""))
                    for c in row
                ):
                    continue
                if "INITIAL STATE" in first or re.match(
                    r"PO[ČC]ETNO\s+STANJE", first
                ):
                    continue

                # Turnover / new balance summary rows
                if first in ("SALES:", "PROMET:"):
                    if len(row) > 4 and row[4]:
                        stmt.total_debit = self.parse_amount_us(row[4])
                    if len(row) > 5 and row[5]:
                        stmt.total_credit = self.parse_amount_us(row[5])
                    continue

                if "NEW BALANCE" in first or "NOVO STANJE" in first:
                    # Closing balance is in the last non-empty cell
                    for c in reversed(row):
                        if c and c.strip():
                            stmt.closing_balance = self.parse_amount_us(c)
                            break
                    continue

                # Skip non-transaction rows (footer, etc.)
                if not re.match(r"\d{2}\.\d{2}\.\d{4}", first):
                    continue

                # Transaction row
                txn = ParsedTransaction(row_number=len(stmt.transactions) + 1)

                # Column 0: Date
                txn.booking_date = self.parse_date_dmy(first)
                txn.value_date = txn.booking_date

                # Column 1: Transaction description (multi-line)
                # Contains: purpose/type on first line, reference on second line
                desc = (row[1] or "").strip() if len(row) > 1 else ""

                # Column 2: Counterparty name + account (multi-line)
                cp_cell = (row[2] or "").strip() if len(row) > 2 else ""

                # Column 3: Additional references (payment code, model numbers)
                ref_cell = (row[3] or "").strip() if len(row) > 3 else ""

                # Parse counterparty
                if cp_cell:
                    cp_lines = [l.strip() for l in cp_cell.split("\n") if l.strip()]
                    if cp_lines:
                        txn.counterparty = cp_lines[0]
                        # Second line is usually the account number
                        if len(cp_lines) > 1:
                            acct = cp_lines[1]
                            if re.match(r"\d{15,18}$", acct):
                                txn.counterparty_account = acct

                # Parse description
                if desc:
                    desc_lines = [l.strip() for l in desc.split("\n") if l.strip()]
                    purpose_parts = []
                    reference = None
                    for dl in desc_lines:
                        # Reference lines look like "0100014600017212 445"
                        if re.match(r"\d{10,}\s+\d+$", dl):
                            reference = dl
                        elif re.match(r"\d{3}\s+", dl):
                            # Payment code prefix like "132 UPLATA..."
                            m = re.match(r"(\d{3})\s+(.+)", dl)
                            if m:
                                txn.payment_code = m.group(1)
                                purpose_parts.append(m.group(2))
                        else:
                            purpose_parts.append(dl)

                    if purpose_parts:
                        txn.purpose = self.clean_text(" ".join(purpose_parts))
                    if reference:
                        txn.reference_debit = reference

                # Parse additional reference
                if ref_cell:
                    ref_lines = [l.strip() for l in ref_cell.split("\n") if l.strip()]
                    if ref_lines:
                        txn.reference_credit = self.clean_text(" ".join(ref_lines))

                # Column 4: CHARGED (debit)
                if len(row) > 4 and row[4]:
                    txn.debit = self.parse_amount_us(row[4])

                # Column 5: IN BENEFIT (credit)
                if len(row) > 5 and row[5]:
                    txn.credit = self.parse_amount_us(row[5])

                if txn.debit is not None or txn.credit is not None:
                    stmt.transactions.append(txn)
