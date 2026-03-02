import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


@register_parser
class ZapadParser(BankParser):
    """Parser for Zapad Banka (570) PDF statements.

    Supports two sub-formats:
    1. Daily statement (Montenegrin): "IZVOD RACUNA - broj N"
       - Old variant: US amounts (1,234.56)
       - New variant: EU amounts with space thousands (1 234,56)
    2. Period statement (English): "ACCOUNT STATEMENT"
    """

    bank_code = "570"
    bank_name = "Zapad Banka"

    def parse(self, file_path: Path) -> ParsedStatement:
        stmt = ParsedStatement(
            bank_code=self.bank_code,
            bank_name=self.bank_name,
        )

        with pdfplumber.open(file_path) as pdf:
            first_text = (pdf.pages[0].extract_text() or "")[:500]
            if "ACCOUNT STATEMENT" in first_text:
                self._parse_period(pdf, stmt)
            else:
                self._parse_daily(pdf, stmt)

        return stmt

    # ----------------------------------------------------------------
    # Helpers
    # ----------------------------------------------------------------
    def _parse_amt(self, text: str, eu_fmt: bool) -> Optional[Decimal]:
        """Parse amount in detected format, stripping space thousands."""
        if not text:
            return None
        text = text.strip().replace(" ", "")
        if eu_fmt:
            return self.parse_amount_eu(text)
        return self.parse_amount_us(text)

    # ----------------------------------------------------------------
    # Daily statement format (Montenegrin)
    # ----------------------------------------------------------------
    def _parse_daily(self, pdf, stmt: ParsedStatement) -> None:
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"

        # Detect amount format: EU (3 546,24) vs US (3,546.24)
        # EU amounts end with ,NN before whitespace; US has ,NNN.NN
        eu_fmt = bool(re.search(r"stanje:\s*\d[\d ]*,\d{2}(?:\s|$)", full_text))

        self._parse_daily_header(full_text, stmt, eu_fmt)
        self._parse_daily_transactions(full_text, stmt, eu_fmt)

    def _parse_daily_header(self, text: str, stmt: ParsedStatement,
                            eu_fmt: bool) -> None:
        # Statement number: "IZVOD RAČUNA - broj 2"
        m = re.search(r"IZVOD\s+RAČUNA\s*-\s*broj\s+(\d+)", text)
        if m:
            stmt.statement_number = m.group(1)

        # Statement date: "za dan 03.01.2026."
        m = re.search(r"za\s+dan\s+(\d{2}\.\d{2}\.\d{4})", text)
        if m:
            stmt.statement_date = self.parse_date_dmy(m.group(1))

        # Client name
        m = re.search(r"Klijent:\s*(.+?)(?:\s{2,}|Žiro)", text)
        if m:
            stmt.client_name = self.clean_text(m.group(1))

        # PIB
        m = re.search(r"JMBG/PIB:\s*(\d+)", text)
        if m:
            stmt.client_pib = m.group(1)

        # Account: "Žiro račun: 570-1110011238-10"
        m = re.search(r"Žiro\s+račun:\s*([\d\-]+)", text)
        if m:
            stmt.account_number = m.group(1)

        # Currency
        m = re.search(r"Valuta:\s*\d+\s+(\w+)", text)
        if m:
            stmt.currency = m.group(1)

        # Amount pattern depends on format
        if eu_fmt:
            ap = r"(\d[\d ]*,\d{2})"
        else:
            ap = r"([\d,]+\.\d{2})"

        m = re.search(rf"Prethodno\s+stanje:\s*{ap}", text)
        if m:
            stmt.opening_balance = self._parse_amt(m.group(1), eu_fmt)

        m = re.search(rf"Krajnje\s+stanje:\s*{ap}", text)
        if m:
            stmt.closing_balance = self._parse_amt(m.group(1), eu_fmt)

        m = re.search(rf"Ukupni\s+promet\s*-\s*duguje:\s*{ap}", text)
        if m:
            stmt.total_debit = self._parse_amt(m.group(1), eu_fmt)
        m = re.search(rf"Ukupni\s+promet\s*-\s*potražuje:\s*{ap}", text)
        if m:
            stmt.total_credit = self._parse_amt(m.group(1), eu_fmt)

    def _parse_daily_transactions(self, text: str, stmt: ParsedStatement,
                                  eu_fmt: bool) -> None:
        """Parse daily format transactions.

        Transaction line: "N. <trans_number> <counterparty> <account> <amounts>"
        Purpose lines follow below, optionally starting with 3-digit payment code.
        """
        lines = text.split("\n")

        # Regex to find EU or US amounts in the "rest" part after account
        if eu_fmt:
            amt_re = re.compile(r"\d{1,3}(?: \d{3})*,\d{2}")
        else:
            amt_re = re.compile(r"[\d,]+\.\d{2}")

        # Transaction start: "N. <trans_no> <counterparty> <account> <rest>"
        txn_re = re.compile(
            r"^(\d+)\.\s+"             # Row number
            r"(\d+)\s+"               # Transaction number
            r"(.+?)\s+"              # Counterparty (non-greedy)
            r"(\d{3}-[\d]+-\d{2})\s+"  # Account (NNN-N...N-NN)
            r"(.+)$"                  # Rest (amounts)
        )

        txn_data = []  # (ParsedTransaction, saldo, needs_direction)

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            m = txn_re.match(line)
            if not m:
                i += 1
                continue

            # Extract amounts from rest of line
            rest = m.group(5)
            amounts_str = amt_re.findall(rest)
            if not amounts_str:
                i += 1
                continue

            amounts = [self._parse_amt(a, eu_fmt) for a in amounts_str]
            amounts = [a for a in amounts if a is not None]
            if not amounts:
                i += 1
                continue

            txn = ParsedTransaction(row_number=int(m.group(1)))
            txn.counterparty = self.clean_text(m.group(3))
            txn.counterparty_account = m.group(4)
            txn.value_date = stmt.statement_date
            txn.booking_date = stmt.statement_date

            if len(amounts) >= 3:
                txn.debit = amounts[0]
                txn.credit = amounts[1]
                saldo = amounts[2]
                needs_direction = False
            elif len(amounts) == 2:
                txn.debit = amounts[0]  # temporary; direction determined later
                txn.credit = Decimal("0")
                saldo = amounts[1]
                needs_direction = True
            else:
                txn.debit = amounts[0]
                txn.credit = Decimal("0")
                saldo = None
                needs_direction = True

            # Collect purpose lines & extract payment code
            purpose_lines = []
            payment_code = None
            i += 1
            while i < len(lines):
                next_line = lines[i].strip()
                if not next_line:
                    i += 1
                    continue
                if (
                    txn_re.match(next_line)
                    or next_line.startswith("UKUPNO:")
                    or "Prethodno stanje:" in next_line
                    or "Krajnje stanje:" in next_line
                    or next_line.startswith("Ovaj dokument")
                    or next_line.startswith("Rbr.")
                    or next_line.startswith("Šifra")
                    or re.match(r"^Zapad banka AD\s*\(", next_line)
                ):
                    break
                # First purpose line may start with 3-digit payment code
                if not purpose_lines:
                    m_code = re.match(r"^(\d{3})\s+(.*)", next_line)
                    if m_code:
                        payment_code = m_code.group(1)
                        rest_text = m_code.group(2).strip()
                        if rest_text:
                            purpose_lines.append(rest_text)
                    else:
                        purpose_lines.append(next_line)
                else:
                    purpose_lines.append(next_line)
                i += 1

            txn.payment_code = payment_code
            if purpose_lines:
                txn.purpose = self.clean_text(" ".join(purpose_lines))

            # Auto-detect bank fees (no payment code, "Fee for order" purpose)
            if not txn.payment_code and txn.purpose and "Fee for order" in txn.purpose:
                txn.payment_code = "221"

            txn_data.append((txn, saldo, needs_direction))

        # Determine debit/credit direction using running balance
        if stmt.opening_balance is not None and txn_data:
            running = stmt.opening_balance
            for txn, saldo, needs_direction in txn_data:
                if needs_direction and saldo is not None:
                    amt = txn.debit or Decimal("0")
                    if running - amt == saldo:
                        txn.debit = amt
                        txn.credit = Decimal("0")
                    elif running + amt == saldo:
                        txn.credit = amt
                        txn.debit = Decimal("0")
                if saldo is not None:
                    running = saldo
        else:
            # Fallback: use totals
            if (stmt.total_credit is not None
                    and stmt.total_credit == Decimal("0")):
                for txn, _, needs_dir in txn_data:
                    if needs_dir:
                        txn.credit = Decimal("0")
            elif (stmt.total_debit is not None
                    and stmt.total_debit == Decimal("0")):
                for txn, _, needs_dir in txn_data:
                    if needs_dir:
                        txn.credit = txn.debit
                        txn.debit = Decimal("0")

        for txn, _, _ in txn_data:
            stmt.transactions.append(txn)

    # ----------------------------------------------------------------
    # Period statement format (English)
    # ----------------------------------------------------------------
    def _parse_period(self, pdf, stmt: ParsedStatement) -> None:
        first_text = pdf.pages[0].extract_text() or ""
        self._parse_period_header(first_text, stmt)

        last_text = pdf.pages[-1].extract_text() or ""
        # Get totals and outgoing balance from last page
        m = re.search(r"OUTGOING\s+BALANCE:\s*([\d,]+\.\d{2})", last_text)
        if m:
            stmt.closing_balance = self.parse_amount_us(m.group(1))

        m = re.search(
            r"TOTAL\s+TURNOVER\s+(?:EUR\(?\d*\)?:?)?\s*([\d,]+\.\d{2})\s+([\d,]+\.\d{2})",
            last_text,
        )
        if m:
            stmt.total_debit = self.parse_amount_us(m.group(1))
            stmt.total_credit = self.parse_amount_us(m.group(2))

        for page in pdf.pages:
            self._parse_period_transactions(page, stmt)

    def _parse_period_header(self, text: str, stmt: ParsedStatement) -> None:
        # Client name: first lines after header
        m = re.search(r"ACCOUNT\s+STATEMENT\s*\n(.+?)(?:\n|ACCOUNT)", text, re.DOTALL)
        if not m:
            # Try from JMBG line area
            m = re.search(r"^(.+?)\n.*?JMBG", text, re.MULTILINE)

        # Client name: on the line after "ACCOUNT STATEMENT"
        # Format: "ROMAX TRADING DOO ACCOUNT PERIOD"
        # The name is before "ACCOUNT" on this line
        lines = text.split("\n")
        for i, line in enumerate(lines):
            if "ACCOUNT STATEMENT" in line:
                # Next line has client name + "ACCOUNT PERIOD"
                if i + 1 < len(lines):
                    name_line = lines[i + 1].strip()
                    # Remove trailing "ACCOUNT PERIOD" or "ACCOUNT" keywords
                    name_line = re.sub(r"\s+ACCOUNT\s+PERIOD\s*$", "", name_line)
                    name_line = re.sub(r"\s+ACCOUNT\s*$", "", name_line)
                    name_line = re.sub(r"\s+PERIOD\s*$", "", name_line)
                    if name_line:
                        stmt.client_name = self.clean_text(name_line)
                break

        # IBAN
        m = re.search(r"IBAN:\s*(ME[\d\s]+)", text)
        if m:
            stmt.iban = m.group(1).replace(" ", "")
            # Derive account number from IBAN
            stmt.account_number = stmt.iban[4:] if len(stmt.iban) > 4 else stmt.iban

        # PIB
        m = re.search(r"JMBG/PIB:\s*(\d+)", text)
        if m:
            stmt.client_pib = m.group(1)

        # Period dates: FROM: DD/MM/YYYY TO: DD/MM/YYYY
        m = re.search(r"FROM:\s*(\d{2}/\d{2}/\d{4})", text)
        if m:
            stmt.period_start = self.parse_date_dmy_slash(m.group(1))
        m = re.search(r"TO:\s*(\d{2}/\d{2}/\d{4})", text)
        if m:
            stmt.period_end = self.parse_date_dmy_slash(m.group(1))
            stmt.statement_date = stmt.period_end

        # Incoming balance
        m = re.search(r"INCOMING\s+BALANCE:\s*([\d,]+\.\d{2})", text)
        if m:
            stmt.opening_balance = self.parse_amount_us(m.group(1))

        # Currency
        m = re.search(r"CURRENCY:\s*(\w+)\s*\((\d+)\)", text)
        if m:
            stmt.currency = m.group(1)

    def _parse_period_transactions(self, page, stmt: ParsedStatement) -> None:
        """Parse period format transactions using text extraction.

        Each transaction block spans multiple lines:
        Line 1 (date area):    DETAILS: <description text>
        <date>                 <more description>
                               <amounts: debit credit balance>
                               <more description>
        <date>  <trans_no>  <recipient/sender>  IBAN: <iban>
        """
        text = page.extract_text() or ""
        lines = text.split("\n")

        i = 0
        while i < len(lines):
            line = lines[i].strip()

            # Look for DETAILS: line which starts a transaction block
            if not line.startswith("DETAILS:") and "DETAILS:" not in line:
                i += 1
                continue

            # Found a DETAILS: line - start collecting transaction block
            block_lines = [line]
            i += 1

            # Collect lines until next DETAILS: or end markers
            while i < len(lines):
                next_line = lines[i].strip()
                if (
                    next_line.startswith("DETAILS:")
                    or "DETAILS:" in next_line
                    or next_line.startswith("TOTAL TURNOVER")
                    or next_line.startswith("OUTGOING BALANCE")
                    or next_line.startswith("This document")
                    or re.match(r"^\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}", next_line)  # timestamp
                ):
                    break
                block_lines.append(next_line)
                i += 1

            self._parse_period_block(block_lines, stmt)

    def _parse_period_block(self, block_lines: list[str], stmt: ParsedStatement) -> None:
        """Parse a single transaction block from period format."""
        block = "\n".join(block_lines)

        # Extract details text
        m = re.search(r"DETAILS:\s*(.+?)(?:\n|$)", block)
        details_text = m.group(1).strip() if m else ""

        # Extract amounts: find line with pattern like "0.30 0.00 420.39" or "315.00 315.00 420.69"
        amounts_pattern = re.compile(
            r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})"
        )
        debit = None
        credit = None
        for line in block_lines:
            m = amounts_pattern.search(line)
            if m:
                debit = self.parse_amount_us(m.group(1))
                credit = self.parse_amount_us(m.group(2))
                break

        if debit is None and credit is None:
            return

        # Extract dates (DD/MM/YYYY)
        dates = re.findall(r"\b(\d{2}/\d{2}/\d{4})\b", block)
        value_date = self.parse_date_dmy_slash(dates[0]) if dates else None
        booking_date = self.parse_date_dmy_slash(dates[1]) if len(dates) > 1 else value_date

        # Extract transaction number
        trans_no = None
        m = re.search(r"\b(\d{7,8})\b", block)
        if m:
            trans_no = m.group(1)

        # Extract IBAN
        iban = None
        m = re.search(r"IBAN:\s*(\S+)", block)
        if m:
            iban = m.group(1)

        # Extract recipient/sender name (on the line with transaction number)
        # Format: "<date> <trans_no> <counterparty name> IBAN: <iban>"
        # For FEE lines without counterparty: "<trans_no> IBAN: <iban> <amounts>"
        counterparty = None
        if trans_no:
            for line in block_lines:
                if trans_no not in line:
                    continue
                # Try to extract name between trans_no and IBAN:
                m = re.search(
                    re.escape(trans_no) + r"\s+(.+?)\s+IBAN:", line
                )
                if m:
                    name = self.clean_text(m.group(1))
                    # Reject if it looks like a number/IBAN (no alpha chars)
                    if name and re.search(r"[A-Za-z]", name):
                        counterparty = name
                break

        # Build purpose from details text plus continuation lines
        purpose_parts = [details_text]
        for line in block_lines[1:]:
            stripped = line.strip()
            # Skip date lines, amount lines, IBAN lines, transaction number lines
            if (
                re.match(r"^\d{2}/\d{2}/\d{4}", stripped)
                or amounts_pattern.search(stripped)
                or "IBAN:" in stripped
                or (trans_no and trans_no in stripped)
                or not stripped
            ):
                continue
            purpose_parts.append(stripped)
        purpose = self.clean_text(" ".join(purpose_parts))

        txn = ParsedTransaction(
            row_number=len(stmt.transactions) + 1,
            value_date=value_date,
            booking_date=booking_date,
            debit=debit,
            credit=credit,
            counterparty=counterparty,
            counterparty_account=iban,
            purpose=purpose,
        )
        stmt.transactions.append(txn)
