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
    # Daily statement format (Montenegrin)
    # ----------------------------------------------------------------
    def _parse_daily(self, pdf, stmt: ParsedStatement) -> None:
        full_text = ""
        for page in pdf.pages:
            full_text += (page.extract_text() or "") + "\n"

        self._parse_daily_header(full_text, stmt)
        self._parse_daily_transactions(full_text, stmt)

    def _parse_daily_header(self, text: str, stmt: ParsedStatement) -> None:
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

        # Prethodno stanje (opening balance)
        m = re.search(r"Prethodno\s+stanje:\s*([\d,]+\.\d{2})", text)
        if m:
            stmt.opening_balance = self.parse_amount_us(m.group(1))

        # Krajnje stanje (closing balance)
        m = re.search(r"Krajnje\s+stanje:\s*([\d,]+\.\d{2})", text)
        if m:
            stmt.closing_balance = self.parse_amount_us(m.group(1))

        # Ukupni promet - duguje / potrazuje
        m = re.search(r"Ukupni\s+promet\s*-\s*duguje:\s*([\d,]+\.\d{2})", text)
        if m:
            stmt.total_debit = self.parse_amount_us(m.group(1))
        m = re.search(r"Ukupni\s+promet\s*-\s*potražuje:\s*([\d,]+\.\d{2})", text)
        if m:
            stmt.total_credit = self.parse_amount_us(m.group(1))

    def _parse_daily_transactions(self, text: str, stmt: ParsedStatement) -> None:
        """Parse daily format transactions from full text.

        Format per transaction block:
        Rbr. Sifra  Broj_trans  Counterparty  Account  Duguje  Potrazuje  Saldo
        Then purpose lines below.

        Lines pattern:
        1. 56218282 Zapad banka AD 570-0000000057001-31 58.80 7,588.45
        OVH SAS\\2, rue Kellermann\\Roubaix\\59100
        """
        lines = text.split("\n")
        # Find transaction lines: start with "N." where N is a digit
        txn_pattern = re.compile(
            r"^\s*(\d+)\.\s+"  # Rbr
            r"(\d+)\s+"  # Broj trans (Sifra field seems missing, just trans number)
            r"(.+?)\s+"  # Counterparty name
            r"(\d{3}-[\d\-]+)\s+"  # Account number (format 570-XXXX-XX)
            r"([\d,]*\.?\d*)\s+"  # Duguje
            r"([\d,]*\.?\d*)\s*$"  # Saldo (Potrazuje may be empty)
        )

        # Alternative: parse more carefully
        # The format is: "1. 56218282 Zapad banka AD 570-0000000057001-31 58.80 7,588.45"
        # After header lines with Rbr/Sifra/Broj trans/etc, transactions follow
        # Each transaction: number. code counterparty account amounts
        # Then purpose lines follow (without leading number)

        i = 0
        row_num = 0
        while i < len(lines):
            line = lines[i].strip()

            # Match transaction start: "N. <code> <name> <account> <amounts>"
            m = re.match(
                r"^(\d+)\.\s+"
                r"(\d+)\s+"  # transaction/sifra number
                r"(.+?)\s+"  # counterparty
                r"(\d{3}-[\d]+(?:-\d+)?)\s+"  # account (NNN-NNNNN-NN format)
                r"([\d,]+\.\d{2})\s+"  # first amount (duguje or potrazuje or saldo)
                r"([\d,]+\.\d{2})\s*$",  # second amount
                line,
            )
            if m:
                row_num += 1
                txn = ParsedTransaction(row_number=int(m.group(1)))
                txn.payment_code = m.group(2)
                txn.counterparty = self.clean_text(m.group(3))
                txn.counterparty_account = m.group(4)

                # Determine debit/credit from amounts
                # Format: Duguje [Potrazuje] Saldo - but in text extraction
                # amounts merge. We get 2 numbers if only debit or credit + saldo.
                amt1 = self.parse_amount_us(m.group(5))
                amt2 = self.parse_amount_us(m.group(6))

                # In the UKUPNO line we see: 58.80 0.00 7,588.45
                # But single transaction line has: 58.80 7,588.45 (debit + saldo)
                # If there are 2 amounts, first is debit, second is saldo
                # We need to figure out if it's debit or credit.
                # Look at the UKUPNO totals: if total credit is 0, all are debits
                txn.debit = amt1
                txn.credit = Decimal("0")

                # Collect purpose from following lines
                purpose_lines = []
                i += 1
                while i < len(lines):
                    next_line = lines[i].strip()
                    if not next_line:
                        i += 1
                        continue
                    # Stop if we hit UKUPNO, next transaction, or known headers
                    if (
                        re.match(r"^\d+\.\s+\d+\s+", next_line)
                        or next_line.startswith("UKUPNO:")
                        or next_line.startswith("Prethodno stanje:")
                        or next_line.startswith("Krajnje stanje:")
                        or next_line.startswith("Ovaj dokument")
                    ):
                        break
                    purpose_lines.append(next_line)
                    i += 1

                if purpose_lines:
                    txn.purpose = self.clean_text(" ".join(purpose_lines))

                txn.value_date = stmt.statement_date
                txn.booking_date = stmt.statement_date
                stmt.transactions.append(txn)
                continue

            # Try 3-amount pattern: "N. code name account debit credit saldo"
            m = re.match(
                r"^(\d+)\.\s+"
                r"(\d+)\s+"
                r"(.+?)\s+"
                r"(\d{3}-[\d]+(?:-\d+)?)\s+"
                r"([\d,]+\.\d{2})\s+"
                r"([\d,]+\.\d{2})\s+"
                r"([\d,]+\.\d{2})\s*$",
                line,
            )
            if m:
                row_num += 1
                txn = ParsedTransaction(row_number=int(m.group(1)))
                txn.payment_code = m.group(2)
                txn.counterparty = self.clean_text(m.group(3))
                txn.counterparty_account = m.group(4)
                txn.debit = self.parse_amount_us(m.group(5))
                txn.credit = self.parse_amount_us(m.group(6))
                txn.value_date = stmt.statement_date
                txn.booking_date = stmt.statement_date

                # Collect purpose
                purpose_lines = []
                i += 1
                while i < len(lines):
                    next_line = lines[i].strip()
                    if not next_line:
                        i += 1
                        continue
                    if (
                        re.match(r"^\d+\.\s+\d+\s+", next_line)
                        or next_line.startswith("UKUPNO:")
                        or next_line.startswith("Prethodno stanje:")
                    ):
                        break
                    purpose_lines.append(next_line)
                    i += 1

                if purpose_lines:
                    txn.purpose = self.clean_text(" ".join(purpose_lines))

                stmt.transactions.append(txn)
                continue

            i += 1

        # Post-process: determine debit vs credit
        # If total_credit is known and is 0, all amounts are debits (already set)
        # If total_debit is known and is 0, all amounts are credits
        if stmt.total_credit == Decimal("0") and stmt.total_debit and stmt.total_debit > 0:
            pass  # debits are correct
        elif stmt.total_debit == Decimal("0") and stmt.total_credit and stmt.total_credit > 0:
            for txn in stmt.transactions:
                txn.credit = txn.debit
                txn.debit = Decimal("0")

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
