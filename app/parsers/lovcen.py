import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


@register_parser
class LovcenParser(BankParser):
    """Parser for Lovcen Banka (565) PDF statements.

    Layout: 8-column table with lines. Very similar to Hipotekarna (520).
    Columns: Valuta, Counterparty+account, Bank, Debit, Credit,
    Payment code+purpose, References, Reclamation.
    Comma decimal (6.851,10). Date DD.MM.YYYY.
    """

    bank_code = "565"
    bank_name = "Lovcen Banka"

    def parse(self, file_path: Path) -> ParsedStatement:
        stmt = ParsedStatement(
            bank_code=self.bank_code,
            bank_name=self.bank_name,
        )

        with pdfplumber.open(file_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"

            self._parse_header(full_text, stmt)

            for page in pdf.pages:
                self._parse_transactions(page, stmt)

        return stmt

    def _parse_header(self, text: str, stmt: ParsedStatement) -> None:
        # Client name: "Klijent :ASTRASOFT DOO"
        m = re.search(r"Klijent\s*:\s*(.+?)(?:\s+PIB|\s*$)", text, re.MULTILINE)
        if m:
            stmt.client_name = self.clean_text(m.group(1))

        # PIB: "PIB :03339645"
        m = re.search(r"PIB\s*:\s*(\d+)", text)
        if m:
            stmt.client_pib = m.group(1)

        # Account: "Broj računa :565000000002145048"
        m = re.search(r"Broj\s+ra[čc]una\s*:\s*(\d{18})", text)
        if m:
            stmt.account_number = m.group(1)

        # Statement: "IZVOD BR. 3 za dan 19.01.2026"
        m = re.search(r"IZVOD\s+BR\.\s*(\d+)\s+za\s+dan\s+(\d{2}\.\d{2}\.\d{4})", text)
        if m:
            stmt.statement_number = m.group(1).lstrip("0") or "0"
            stmt.statement_date = self.parse_date_dmy(m.group(2))

    def _parse_transactions(self, page, stmt: ParsedStatement) -> None:
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                if not row or not row[0]:
                    continue

                cell0 = row[0].strip()

                # Summary table: "Predhodno stanje" header row
                if cell0 in ("Predhodno stanje", "Prethodno stanje"):
                    continue
                # Summary data row: opening balance in EU format
                if re.match(r"[\d.,]+$", cell0) and len(row) >= 4:
                    # Check if this is the summary row
                    if self._is_summary_row(row):
                        stmt.opening_balance = self.parse_amount_eu(row[0])
                        if len(row) > 1:
                            stmt.total_debit = self.parse_amount_eu(row[1])
                        if len(row) > 2:
                            stmt.total_credit = self.parse_amount_eu(row[2])
                        if len(row) > 3:
                            stmt.closing_balance = self.parse_amount_eu(row[3])
                        continue

                # Skip header rows
                if cell0 in ("1", "Valuta") or "Naziv i" in cell0:
                    continue
                # Skip column number row
                if re.match(r"^\d$", cell0):
                    continue

                # Transaction rows start with a date DD.MM.YYYY
                date_match = re.match(r"(\d{2}\.\d{2}\.\d{4})", cell0)
                if not date_match:
                    continue

                txn = ParsedTransaction(
                    row_number=len(stmt.transactions) + 1,
                )
                txn.value_date = self.parse_date_dmy(date_match.group(1))
                txn.booking_date = txn.value_date

                # Column 1-2: Counterparty name + account (may have None for col 2)
                if len(row) > 1 and row[1]:
                    self._parse_counterparty(row[1], txn)

                # Column 2/3: Bank (may be None if merged)
                if len(row) > 2 and row[2]:
                    txn.counterparty_bank = self.clean_text(row[2])

                # Column 3/4: Debit (Zaduzenje)
                debit_col = 3
                if len(row) > debit_col and row[debit_col]:
                    txn.debit = self.parse_amount_eu(row[debit_col])
                    if txn.debit is not None and txn.debit == Decimal("0"):
                        txn.debit = None

                # Column 4/5: Credit (Odobrenje)
                credit_col = 4
                if len(row) > credit_col and row[credit_col]:
                    txn.credit = self.parse_amount_eu(row[credit_col])
                    if txn.credit is not None and txn.credit == Decimal("0"):
                        txn.credit = None

                # Column 5/6: Payment code + purpose
                purpose_col = 5
                if len(row) > purpose_col and row[purpose_col]:
                    self._parse_purpose(row[purpose_col], txn)

                # Column 6/7: References (Zaduzenje/Odobrenje)
                ref_col = 6
                if len(row) > ref_col and row[ref_col]:
                    txn.reference_debit = self.clean_text(row[ref_col])

                # Column 7/8: Reclamation data
                reclam_col = 7
                if len(row) > reclam_col and row[reclam_col]:
                    txn.reclamation_data = self.clean_text(row[reclam_col])

                if txn.debit is not None or txn.credit is not None:
                    stmt.transactions.append(txn)

    def _is_summary_row(self, row) -> bool:
        """Check if a row is the summary (balance) row."""
        # Summary row has numeric values in most columns followed by integers
        if len(row) < 4:
            return False
        try:
            for i in range(4):
                if row[i] and not re.match(r"^[\d.,\s]+$", row[i].strip()):
                    return False
            return True
        except (IndexError, AttributeError):
            return False

    def _parse_counterparty(self, cell: str, txn: ParsedTransaction) -> None:
        """Parse counterparty cell containing name and account on separate lines."""
        lines = [l.strip() for l in cell.split("\n") if l.strip()]
        if not lines:
            return

        name_parts = []
        for line in lines:
            # 18-digit account number
            if re.match(r"^\d{18}$", line):
                txn.counterparty_account = line
            # Account in XXX-XXXXX-XX format
            elif re.match(r"^\d{3}-\d+-\d{2}$", line):
                txn.counterparty_account = line
            else:
                name_parts.append(line)

        if name_parts:
            txn.counterparty = self.clean_text(" ".join(name_parts))

    def _parse_purpose(self, cell: str, txn: ParsedTransaction) -> None:
        """Parse payment code and purpose from the combined cell."""
        lines = [l.strip() for l in cell.split("\n") if l.strip()]
        if not lines:
            return

        # First line: "163 W06 OSTALI TRANSFERI" or "400 WNK"
        first = lines[0]
        m = re.match(r"(\d{3})\s+(.*)", first)
        if m:
            txn.payment_code = m.group(1)
            purpose_parts = [m.group(2)]
        else:
            purpose_parts = [first]

        # Additional lines are continuation of purpose
        purpose_parts.extend(lines[1:])
        txn.purpose = self.clean_text(" ".join(purpose_parts))
