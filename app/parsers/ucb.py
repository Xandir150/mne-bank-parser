import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


@register_parser
class UCBParser(BankParser):
    """Parser for Universal Capital Bank (560) PDF statements."""

    bank_code = "560"
    bank_name = "Universal Capital Bank"

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

    def _parse_header(self, page, stmt: ParsedStatement) -> None:
        text = page.extract_text() or ""

        # Client name: "Naziv: ASTRASOFT"
        m = re.search(r"Naziv:\s*(.+)", text)
        if m:
            name = m.group(1).strip()
            # Stop at known field boundaries
            for stop in ["Mjesto:", "Matični", "Izvod"]:
                idx = name.find(stop)
                if idx > 0:
                    name = name[:idx].strip()
            stmt.client_name = self.clean_text(name)

        # Account number from "Broj partije: 560-0000000002903-42"
        m = re.search(r"Broj\s+partije:\s*([\d\-]+)", text)
        if m:
            stmt.account_number = m.group(1).strip()

        # Statement number: "Izvod broj : 8"
        m = re.search(r"Izvod\s+broj\s*:\s*(\d+)", text)
        if m:
            stmt.statement_number = m.group(1)

        # Statement date: "STANJE I PROMJENE SREDSTAVA NA DAN DD.MM.YYYY"
        m = re.search(r"NA\s+DAN\s+(\d{2})\.(\d{2})\.(\d{4})", text)
        if m:
            try:
                stmt.statement_date = date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                pass

        # PIB
        m = re.search(r"(?:Poreski\s+broj|PIB):\s*(\d+)", text)
        if m:
            stmt.client_pib = m.group(1)

        # Balances from the summary table
        tables = page.extract_tables()
        if tables:
            # First table is the summary: Prethodno stanje, Duguje, Potrazuje, Novo stanje
            summary = tables[0]
            for row in summary:
                if row and row[0] and re.match(r"[\d,]+\.\d{2}", row[0].strip()):
                    stmt.opening_balance = self.parse_amount_us(row[0])
                    if len(row) > 1:
                        stmt.total_debit = self.parse_amount_us(row[1])
                    if len(row) > 2:
                        stmt.total_credit = self.parse_amount_us(row[2])
                    if len(row) > 3:
                        stmt.closing_balance = self.parse_amount_us(row[3])
                    break

    def _parse_transactions(self, page, stmt: ParsedStatement) -> None:
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                if not row or not row[0]:
                    continue
                rb = row[0].strip()
                # Transaction rows start with a number (RB)
                if not rb.isdigit():
                    # Check for "Ukupno EUR:" totals row
                    if "Ukupno" in rb:
                        if len(row) > 3 and row[3]:
                            stmt.total_debit = self.parse_amount_us(row[3])
                        if len(row) > 4 and row[4]:
                            stmt.total_credit = self.parse_amount_us(row[4])
                    continue

                txn = ParsedTransaction(row_number=int(rb))

                # Column 1: Name + address + account (multiline)
                if len(row) > 1 and row[1]:
                    self._parse_counterparty(row[1], txn)

                # Column 2: Origin + date (e.g. "00-Podgorica/ 2026.02.02")
                if len(row) > 2 and row[2]:
                    m = re.search(r"(\d{4}\.\d{2}\.\d{2})", row[2])
                    if m:
                        txn.booking_date = self.parse_date_ymd(m.group(1))
                        txn.value_date = txn.booking_date

                # Column 3: Debit amount
                if len(row) > 3 and row[3]:
                    txn.debit = self.parse_amount_us(row[3])

                # Column 4: Credit amount
                if len(row) > 4 and row[4]:
                    txn.credit = self.parse_amount_us(row[4])

                # Column 5: Payment code (Sifra)
                if len(row) > 5 and row[5]:
                    txn.payment_code = row[5].strip()

                # Column 6: Purpose
                if len(row) > 6 and row[6]:
                    txn.purpose = self.clean_text(row[6])

                # Column 7: References (odobrenja/zaduzenja)
                if len(row) > 7 and row[7]:
                    ref_text = row[7].strip()
                    if ref_text:
                        txn.reference_credit = self.clean_text(ref_text)

                # Column 8: Reclamation data
                if len(row) > 8 and row[8]:
                    txn.reclamation_data = self.clean_text(row[8])

                # Only add if we have meaningful data
                if txn.debit is not None or txn.credit is not None:
                    stmt.transactions.append(txn)

    def _parse_counterparty(self, cell: str, txn: ParsedTransaction) -> None:
        """Parse counterparty cell: name, address, account number."""
        lines = [l.strip() for l in cell.split("\n") if l.strip()]
        if not lines:
            return

        # Account number is the last long digit string (15-18 digits)
        account = None
        name_parts = []
        for line in lines:
            # Check if line ends with or is an account number
            m = re.search(r"(\d{15,18})\s*$", line)
            if m:
                account = m.group(1)
                prefix = line[: m.start()].strip().rstrip(",").strip()
                if prefix:
                    name_parts.append(prefix)
            else:
                name_parts.append(line)

        txn.counterparty_account = account
        full = ", ".join(name_parts)
        # Clean up double commas and spaces
        full = re.sub(r",\s*,", ",", full)
        full = self.clean_text(full)
        txn.counterparty = full
