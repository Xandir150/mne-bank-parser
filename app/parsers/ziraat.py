import re
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


@register_parser
class ZiraatParser(BankParser):
    """Parser for Ziraat Bank Montenegro (575) PDF statements.

    Similar structure to Prva Banka: table with RB-numbered rows,
    each having a Naknada (fee) sub-line within the Zaduzenje cell.
    Counterparty names are often concatenated without spaces.
    """

    bank_code = "575"
    bank_name = "Ziraat Bank Montenegro"

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

        # Client name: "Naziv: BALKAN ART DOO BUDVA"
        m = re.search(r"Naziv:\s*(.+?)(?:\s{2,}|Matični|$)", text, re.MULTILINE)
        if m:
            stmt.client_name = self.clean_text(m.group(1))

        # Account: "Račun: 575-0000000002269-08"
        m = re.search(r"Račun:\s*([\d\-]+)", text)
        if m:
            stmt.account_number = m.group(1)

        # Statement number: "IZVOD BROJ 230"
        m = re.search(r"IZVOD\s+BROJ\s+(\d+)", text)
        if m:
            stmt.statement_number = m.group(1)

        # Statement date: "STANJE I PROMJENE SREDSTAVA NA DAN 01.12.2025"
        m = re.search(r"NA\s+DAN\s+(\d{2}\.\d{2}\.\d{4})", text)
        if m:
            stmt.statement_date = self.parse_date_dmy(m.group(1))

        # PIB
        m = re.search(r"PIB:\s*(\d+)", text)
        if m:
            stmt.client_pib = m.group(1)

        # Balances from summary table
        tables = page.extract_tables()
        if tables:
            for table in tables:
                for row in table:
                    if not row or not row[0]:
                        continue
                    val = row[0].strip() if row[0] else ""
                    # Look for the data row with opening balance
                    if re.match(r"[\d,]+\.\d{2}$", val):
                        stmt.opening_balance = self.parse_amount_us(val)
                        if len(row) > 1 and row[1]:
                            stmt.total_debit = self.parse_amount_us(row[1])
                        if len(row) > 2 and row[2]:
                            stmt.total_credit = self.parse_amount_us(row[2])
                        if len(row) > 3 and row[3]:
                            stmt.closing_balance = self.parse_amount_us(row[3])
                        break

    def _parse_transactions(self, page, stmt: ParsedStatement) -> None:
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                if not row or not row[0]:
                    continue
                rb = row[0].strip()
                if not rb.isdigit():
                    # Check for UKUPNO totals
                    if "UKUPNO" in (row[0] or ""):
                        # Parse totals - format: "UKUPNO:\nNaknada:"
                        for ci in range(len(row)):
                            cell = (row[ci] or "").strip()
                            if ci >= 3 and cell:
                                lines = cell.split("\n")
                                amt = self.parse_amount_us(lines[0]) if lines else None
                                if ci == 3 and amt:
                                    stmt.total_debit = amt
                                elif ci == 4 and amt:
                                    stmt.total_credit = amt
                    continue

                txn = ParsedTransaction(row_number=int(rb))

                # Column 1: Counterparty + account (names often concatenated)
                if len(row) > 1 and row[1]:
                    self._parse_counterparty(row[1], txn)

                # Column 2: Origin + date
                if len(row) > 2 and row[2]:
                    m = re.search(r"(\d{2}\.\d{2}\.\d{4})", row[2])
                    if m:
                        txn.booking_date = self.parse_date_dmy(m.group(1))
                        txn.value_date = txn.booking_date

                # Column 3: Debit amount + Naknada (fee) on second line
                if len(row) > 3 and row[3]:
                    self._parse_debit_with_fee(row[3], txn)

                # Column 4: Credit amount
                if len(row) > 4 and row[4]:
                    txn.credit = self.parse_amount_us(row[4])

                # Column 5: Payment code (Sifra)
                if len(row) > 5 and row[5]:
                    txn.payment_code = row[5].strip()

                # Column 6: Purpose (Svrha placanja)
                if len(row) > 6 and row[6]:
                    txn.purpose = self.clean_text(row[6])

                # Column 7: References (model Broj zaduzenja / odobrenja)
                if len(row) > 7 and row[7]:
                    ref = row[7].strip()
                    if ref and ref != "( )\n( )":
                        # Split into debit/credit references
                        ref_lines = [l.strip() for l in ref.split("\n") if l.strip()]
                        if ref_lines:
                            txn.reference_debit = ref_lines[0] if ref_lines[0] != "( )" else None
                        if len(ref_lines) > 1:
                            txn.reference_credit = (
                                ref_lines[1] if ref_lines[1] != "( )" else None
                            )

                # Column 8: Reclamation data
                if len(row) > 8 and row[8]:
                    txn.reclamation_data = self.clean_text(row[8])

                if txn.debit is not None or txn.credit is not None:
                    stmt.transactions.append(txn)

    def _parse_counterparty(self, cell: str, txn: ParsedTransaction) -> None:
        """Parse counterparty cell. Names are often concatenated without spaces."""
        lines = [l.strip() for l in cell.split("\n") if l.strip()]
        if not lines:
            return

        # Last line usually contains the account number (format NNN-XXXXX-NN or NNN-N-NN)
        account = None
        name_lines = []
        for line in lines:
            m = re.match(r"^(\d{3}-[\d]+-\d{2})$", line)
            if m:
                account = m.group(1)
            else:
                name_lines.append(line)

        txn.counterparty_account = account
        txn.counterparty = self.clean_text(", ".join(name_lines)) if name_lines else None

    def _parse_debit_with_fee(self, cell: str, txn: ParsedTransaction) -> None:
        """Parse debit cell that contains amount + Naknada (fee) on separate lines.

        Example: "1,328.85\nNaknada 1.75" or "162.03\nNaknada 0.00"
        """
        lines = [l.strip() for l in cell.split("\n") if l.strip()]
        if not lines:
            return

        txn.debit = self.parse_amount_us(lines[0])

        if len(lines) > 1:
            m = re.search(r"Naknada\s+([\d,]+\.\d{2})", lines[1])
            if m:
                fee = self.parse_amount_us(m.group(1))
                if fee and fee > Decimal("0"):
                    txn.fee = fee
