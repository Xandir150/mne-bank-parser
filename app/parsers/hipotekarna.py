import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


@register_parser
class HipotekarnaParser(BankParser):
    """Parser for Hipotekarna Banka (520) PDF statements.

    Layout: text-only PDF (no vector lines). Header has client/bank info,
    then IZVOD BR. line, then transaction rows, then summary.
    8-column table: Valuta, Counterparty+account, Bank, Debit, Credit,
    Payment code+purpose, References, Reclamation.
    Period decimal (1069.94). Date DD.MM.YYYY.
    """

    bank_code = "520"
    bank_name = "Hipotekarna Banka"

    def parse(self, file_path: Path) -> ParsedStatement:
        stmt = ParsedStatement(
            bank_code=self.bank_code,
            bank_name=self.bank_name,
        )

        with pdfplumber.open(file_path) as pdf:
            all_words = []
            for page in pdf.pages:
                words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
                all_words.extend(words)

            first_page_words = pdf.pages[0].extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
            self._parse_header(first_page_words, stmt)

            for page in pdf.pages:
                self._parse_page(page, stmt)

        return stmt

    def _group_by_y(self, words, tolerance=5):
        """Group words into lines by y-position."""
        from collections import defaultdict
        y_groups = defaultdict(list)
        for w in words:
            y_key = round(w['top'] / tolerance) * tolerance
            y_groups[y_key].append(w)
        result = []
        for y_key in sorted(y_groups.keys()):
            ws = sorted(y_groups[y_key], key=lambda w: w['x0'])
            result.append((y_key, ws))
        return result

    def _parse_header(self, words: list, stmt: ParsedStatement) -> None:
        lines = self._group_by_y(words)

        # Header structure (from visual):
        # y~45: client_name (x~998) | PIB (x~1528)
        # y~70: place
        # y~100: address | currency code (x~1528)
        # y~125: account number (18 digits)
        # y~210: statement_number | statement_date
        for y, ws in lines:
            texts = [(w['x0'], w['text'].strip()) for w in ws if w['text'].strip()]
            if not texts:
                continue

            for x, t in texts:
                # PIB is a short numeric at right side
                if re.match(r"^\d{7,8}$", t) and x > 1000:
                    stmt.client_pib = t
                # Currency code
                elif re.match(r"^\d{3}$", t) and x > 1000:
                    pass  # currency code like 978
                # 18-digit account number
                elif re.match(r"^(520|565)\d{15}$", t):
                    stmt.account_number = t
                # Statement number + date line: "004" at one x, "01.02.2026." at another
                elif re.match(r"^\d{1,4}$", t) and not stmt.statement_number:
                    # Check if there's a date on the same line
                    for x2, t2 in texts:
                        if re.match(r"\d{2}\.\d{2}\.\d{4}", t2):
                            stmt.statement_number = t.lstrip("0") or "0"
                            stmt.statement_date = self.parse_date_dmy(t2)
                            break

            # Client name: first text at x~998 (right header block), not a number
            if not stmt.client_name:
                for x, t in texts:
                    if x > 900 and not re.match(r"^[\d.]+$", t) and len(t) > 3:
                        stmt.client_name = self.clean_text(t)
                        break

    def _parse_page(self, page, stmt: ParsedStatement) -> None:
        words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
        lines = self._group_by_y(words)
        text = page.extract_text() or ""

        # Parse transactions from text lines
        # Transaction line: date + counterparty_bank + debit + credit
        # Next line: account + purpose + reclamation
        all_lines = text.split("\n")

        i = 0
        while i < len(all_lines):
            line = all_lines[i].strip()

            # Match transaction start: DD.MM.YYYY. followed by text and amounts
            m = re.match(
                r"(\d{2}\.\d{2}\.\d{4})\.?\s+(.+?)\s+([\d,.]+)\s+([\d,.]+)\s*$",
                line,
            )
            if m:
                date_str = m.group(1)
                counterparty_bank = m.group(2).strip()
                debit_str = m.group(3)
                credit_str = m.group(4)

                txn = ParsedTransaction(
                    row_number=len(stmt.transactions) + 1,
                )
                txn.value_date = self.parse_date_dmy(date_str)
                txn.booking_date = txn.value_date
                txn.counterparty_bank = self.clean_text(counterparty_bank)

                txn.debit = self.parse_amount_us(debit_str)
                txn.credit = self.parse_amount_us(credit_str)
                if txn.debit is not None and txn.debit == Decimal("0"):
                    txn.debit = None
                if txn.credit is not None and txn.credit == Decimal("0"):
                    txn.credit = None

                # Next line: account + purpose + reclamation
                if i + 1 < len(all_lines):
                    next_line = all_lines[i + 1].strip()
                    m2 = re.match(r"(\d{18})\s*(.*)", next_line)
                    if m2:
                        txn.counterparty_account = m2.group(1)
                        rest = m2.group(2).strip()
                        # Split purpose and reclamation reference
                        m3 = re.search(r"(\d{3}-\d{9,15})\s*$", rest)
                        if m3:
                            txn.purpose = self.clean_text(rest[: m3.start()])
                            txn.reclamation_data = m3.group(1)
                        elif rest:
                            txn.purpose = self.clean_text(rest)
                        i += 1

                if txn.debit is not None or txn.credit is not None:
                    stmt.transactions.append(txn)
            else:
                # Summary line: opening debit credit closing count_debit count_credit
                m = re.match(
                    r"([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+)\s+(\d+)\s*$",
                    line,
                )
                if m:
                    stmt.opening_balance = self.parse_amount_us(m.group(1))
                    stmt.total_debit = self.parse_amount_us(m.group(2))
                    stmt.total_credit = self.parse_amount_us(m.group(3))
                    stmt.closing_balance = self.parse_amount_us(m.group(4))

            i += 1
