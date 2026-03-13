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
                # 18-digit account number (only set once from header area)
                elif re.match(r"^(520|565)\d{15}$", t) and not stmt.account_number:
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

    # Column x-boundaries (from PDF word positions):
    # x < 150: date (col 1)
    # 150 <= x < 600: counterparty name / account (col 2)
    # 600 <= x < 800: debit amount (col 4)
    # 800 <= x < 880: credit amount (col 5)
    # 880 <= x < 1200: payment code / purpose (col 6)
    # 1200 <= x < 1450: references (col 7)
    # x >= 1450: reclamation data (col 8)
    _COL_COUNTERPARTY = 150
    _COL_DEBIT = 600
    _COL_CREDIT = 800
    _COL_PURPOSE = 880
    _COL_REF = 1200
    _COL_RECLAM = 1450

    def _classify_word(self, x: float) -> str:
        if x < self._COL_COUNTERPARTY:
            return "date"
        elif x < self._COL_DEBIT:
            return "counterparty"
        elif x < self._COL_CREDIT:
            return "debit"
        elif x < self._COL_PURPOSE:
            return "credit"
        elif x < self._COL_REF:
            return "purpose"
        elif x < self._COL_RECLAM:
            return "reference"
        else:
            return "reclamation"

    def _parse_page(self, page, stmt: ParsedStatement) -> None:
        words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
        lines = self._group_by_y(words, tolerance=4)

        # Pair lines: line1 (date row) + line2 (account row)
        pending_txn = None

        for y, ws in lines:
            cols = {}
            for w in ws:
                text = w['text'].strip()
                if not text:
                    continue
                col = self._classify_word(w['x0'])
                cols.setdefault(col, []).append(text)

            date_text = " ".join(cols.get("date", []))

            # Check if this is a summary line (all numbers, no date pattern)
            if not re.match(r"\d{2}\.\d{2}\.\d{4}", date_text):
                # Could be line 2 of a transaction or summary
                cp_text = " ".join(cols.get("counterparty", []))

                if pending_txn and re.match(r"(\d{18}|\d{3}-\d{5}-\d{2})", cp_text):
                    # Line 2: account, purpose, reference, reclamation
                    pending_txn.counterparty_account = cp_text.split()[0]
                    pending_txn.purpose = self.clean_text(" ".join(cols.get("purpose", [])))
                    ref_text = " ".join(cols.get("reference", []))
                    if ref_text:
                        pending_txn.reference_credit = ref_text
                    reclam_text = " ".join(cols.get("reclamation", []))
                    if reclam_text:
                        pending_txn.reclamation_data = reclam_text
                    if pending_txn.debit is not None or pending_txn.credit is not None:
                        stmt.transactions.append(pending_txn)
                    pending_txn = None
                    continue

                # Summary line
                all_text = " ".join(t for texts in cols.values() for t in texts)
                m = re.match(
                    r"([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+([\d,.]+)\s+(\d+)\s+(\d+)\s*$",
                    all_text,
                )
                if m:
                    stmt.opening_balance = self.parse_amount_us(m.group(1))
                    stmt.total_debit = self.parse_amount_us(m.group(2))
                    stmt.total_credit = self.parse_amount_us(m.group(3))
                    stmt.closing_balance = self.parse_amount_us(m.group(4))
                continue

            # Line 1: date, counterparty, debit, credit, payment_code, reference
            if pending_txn and (pending_txn.debit is not None or pending_txn.credit is not None):
                # Previous txn had no line 2 (shouldn't happen, but safe)
                stmt.transactions.append(pending_txn)
                pending_txn = None

            date_str = re.match(r"(\d{2}\.\d{2}\.\d{4})", date_text).group(1)
            counterparty_name = " ".join(cols.get("counterparty", []))
            debit_str = " ".join(cols.get("debit", []))
            credit_str = " ".join(cols.get("credit", []))
            purpose_parts = cols.get("purpose", [])
            payment_code = purpose_parts[0] if purpose_parts else ""
            ref_debit = " ".join(cols.get("reference", []))
            reclam = " ".join(cols.get("reclamation", []))

            txn = ParsedTransaction(
                row_number=len(stmt.transactions) + 1,
            )
            txn.value_date = self.parse_date_dmy(date_str)
            txn.booking_date = txn.value_date
            txn.counterparty = self.clean_text(counterparty_name)
            txn.payment_code = payment_code if re.match(r"\d{3}$", payment_code) else ""

            txn.debit = self.parse_amount_us(debit_str) if debit_str else None
            txn.credit = self.parse_amount_us(credit_str) if credit_str else None
            if txn.debit is not None and txn.debit == Decimal("0"):
                txn.debit = None
            if txn.credit is not None and txn.credit == Decimal("0"):
                txn.credit = None

            if ref_debit:
                txn.reference_debit = ref_debit
            if reclam:
                txn.reclamation_data = reclam

            pending_txn = txn

        # Flush last pending
        if pending_txn and (pending_txn.debit is not None or pending_txn.credit is not None):
            stmt.transactions.append(pending_txn)
