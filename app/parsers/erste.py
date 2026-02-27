import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

from bs4 import BeautifulSoup

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


@register_parser
class ErsteParser(BankParser):
    """Parser for Erste Bank (540) HTML statements.

    Layout: HTML file with windows-1250 encoding. Nested HTML tables.
    Transaction rows have 3 dates stacked (document/value/processing),
    counterparty info, sequential number + purpose + payment code,
    references, debit, credit.
    Comma decimal (2.256,68). Date DD.MM.YYYY. with trailing dot.
    """

    bank_code = "540"
    bank_name = "Erste Bank"

    def parse(self, file_path: Path) -> ParsedStatement:
        stmt = ParsedStatement(
            bank_code=self.bank_code,
            bank_name=self.bank_name,
        )

        # Read with windows-1250 encoding
        raw = file_path.read_bytes()
        try:
            html = raw.decode("windows-1250")
        except UnicodeDecodeError:
            html = raw.decode("utf-8", errors="replace")

        soup = BeautifulSoup(html, "html.parser")
        self._parse_header(soup, stmt)
        self._parse_transactions(soup, stmt)

        return stmt

    def _parse_header(self, soup: BeautifulSoup, stmt: ParsedStatement) -> None:
        # Date is in a <p> tag: "Za period (po datumu obrade): DD.MM.YYYY."
        for p in soup.find_all("p"):
            text = p.get_text(" ", strip=True)
            m = re.search(r"Za\s+period.*?:\s*(\d{2}\.\d{2}\.\d{4})", text)
            if m and not stmt.statement_date:
                stmt.statement_date = self.parse_date_dmy(m.group(1))

        # Find all tables
        tables = soup.find_all("table")

        for table in tables:
            text = table.get_text(" ", strip=True)

            # Client name: "Naziv klijenta: NAME"
            m = re.search(r"Naziv\s+klijenta:\s*(.+?)(?:\s*$)", text)
            if m and not stmt.client_name:
                stmt.client_name = self.clean_text(m.group(1))

            # Account: "Broj računa: 540-000000001422553"
            m = re.search(r"Broj\s+ra[čc]una:\s*(540[\d-]+)", text)
            if m and not stmt.account_number:
                stmt.account_number = m.group(1)

            # Statement number: "Broj izvoda: 001/2026"
            m = re.search(r"Broj\s+izvoda:\s*(\S+)", text)
            if m and not stmt.statement_number:
                stmt.statement_number = m.group(1)

            # Currency
            m = re.search(r"Oznaka\s+valute:\s*(\w+)", text)
            if m:
                stmt.currency = m.group(1)

            # Client info from STASSI ATHLETICS DOO line
            rows = table.find_all("tr")
            for row in rows:
                cells = row.find_all("td")
                if len(cells) >= 2:
                    value = cells[1].get_text(strip=True)
                    if not stmt.client_name and value and "DOO" in value:
                        stmt.client_name = self.clean_text(value)

    def _parse_transactions(self, soup: BeautifulSoup, stmt: ParsedStatement) -> None:
        # Find the main transaction table (contains "Datum dokumenta" header)
        tables = soup.find_all("table")
        txn_table = None
        for table in tables:
            header_text = table.get_text(" ", strip=True)
            if "Datum dokumenta" in header_text:
                txn_table = table
                break

        if not txn_table:
            return

        rows = txn_table.find_all("tr")

        for row in rows:
            cells = row.find_all("td")
            if not cells:
                continue

            first_cell_text = cells[0].get_text(strip=True)

            # Opening balance row: "Početno stanje" ... amount (3 cells)
            if "stanje" in first_cell_text.lower() and "po" in first_cell_text.lower():
                balance_text = cells[-1].get_text(strip=True)
                stmt.opening_balance = self.parse_amount_eu(balance_text)
                continue

            # Closing balance row: "Konačno stanje" ... amount (3 cells)
            if "stanje" in first_cell_text.lower() and "kon" in first_cell_text.lower():
                balance_text = cells[-1].get_text(strip=True)
                stmt.closing_balance = self.parse_amount_eu(balance_text)
                continue

            # Summary row: "Stanje na dan" with totals (6 cells)
            if "Stanje na dan" in first_cell_text:
                if len(cells) >= 5:
                    # Second-to-last cell: debit total (first line before line break)
                    debit_text = cells[-2].get_text("\n", strip=True).split("\n")[0].strip()
                    # Last cell: credit total + closing balance (first line)
                    credit_text = cells[-1].get_text("\n", strip=True).split("\n")[0].strip()
                    stmt.total_debit = self.parse_amount_eu(debit_text)
                    stmt.total_credit = self.parse_amount_eu(credit_text)
                continue

            # Skip header rows
            if "Datum dokumenta" in first_cell_text or not first_cell_text:
                continue

            # Transaction row: first cell has 3 dates separated by <br>
            dates_match = re.findall(r"(\d{2}\.\d{2}\.\d{4})", first_cell_text)
            if not dates_match or len(cells) < 5:
                continue

            txn = ParsedTransaction(
                row_number=len(stmt.transactions) + 1,
            )

            # Dates: document date, value date, processing date
            if len(dates_match) >= 1:
                txn.booking_date = self.parse_date_dmy(dates_match[0])
            if len(dates_match) >= 2:
                txn.value_date = self.parse_date_dmy(dates_match[1])

            # Column 2: Counterparty info (name, account, exchange rate)
            if len(cells) > 1:
                self._parse_counterparty_cell(cells[1], txn)

            # Column 3: Sequential number + purpose + payment code
            if len(cells) > 2:
                self._parse_purpose_cell(cells[2], txn)

            # Column 4: References (debit ref, credit ref, transaction ref)
            if len(cells) > 3:
                self._parse_references_cell(cells[3], txn)

            # Column 5: Debit (Na teret)
            if len(cells) > 4:
                debit_text = cells[4].get_text(strip=True)
                txn.debit = self.parse_amount_eu(debit_text)
                if txn.debit is not None and txn.debit == Decimal("0"):
                    txn.debit = None

            # Column 6: Credit (U korist)
            if len(cells) > 5:
                credit_text = cells[5].get_text(strip=True)
                txn.credit = self.parse_amount_eu(credit_text)
                if txn.credit is not None and txn.credit == Decimal("0"):
                    txn.credit = None

            if txn.debit is not None or txn.credit is not None:
                stmt.transactions.append(txn)

    def _parse_counterparty_cell(self, cell, txn: ParsedTransaction) -> None:
        """Parse counterparty cell with name, account, and exchange rate."""
        text = cell.get_text("\n", strip=True)
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        if not lines:
            return

        # First line: counterparty name
        txn.counterparty = self.clean_text(lines[0])

        # Second line: account number (540-XXXX or other bank format)
        if len(lines) > 1:
            acct = lines[1].strip()
            if re.match(r"\d{3}-", acct):
                txn.counterparty_account = acct

    def _parse_purpose_cell(self, cell, txn: ParsedTransaction) -> None:
        """Parse purpose cell: sequential number, purpose, payment code."""
        text = cell.get_text("\n", strip=True)
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        if not lines:
            return

        # First line: "1 - PAYPAL *RROZENBERG" or "1 - purpose text"
        first = lines[0]
        m = re.match(r"\d+\s*-\s*(.*)", first)
        if m:
            txn.purpose = self.clean_text(m.group(1))
        else:
            txn.purpose = self.clean_text(first)

        # Payment code might be on the last line or embedded
        # Sometimes there's no separate payment code line

    def _parse_references_cell(self, cell, txn: ParsedTransaction) -> None:
        """Parse references cell with debit ref, credit ref, transaction ref."""
        text = cell.get_text("\n", strip=True)
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        if not lines:
            return

        if len(lines) >= 1:
            txn.reference_debit = self.clean_text(lines[0])
        if len(lines) >= 2:
            txn.reference_credit = self.clean_text(lines[1])
        if len(lines) >= 3:
            txn.reclamation_data = self.clean_text(lines[2])
