import re
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


@register_parser
class CKBParser(BankParser):
    """Parser for CKB / Crnogorska Komercijalna Banka (510) PDF statements.

    Format: "Izvod broj N za promet i stanje računa <18 digits> na dan DD.MM.YYYY"
    Each transaction is a 3-line block:
      Line 1: <Rbr> <ID> <Račun> [Šifra] <date1> <date2> <Odliv|Priliv> <Provizija> [Reference]
      Line 2: counterparty name + city
      Line 3: purpose
    """

    bank_code = "510"
    bank_name = "CKB Banka"

    # Column x0 boundaries derived from layout
    X_ODLIV_MAX = 555  # x0 below this = debit (Odliv)
    X_PRILIV_MAX = 625  # x0 below this (and >= 555) = credit (Priliv)
    X_PROVIZIJA_MAX = 700  # x0 below this (and >= 625) = fee
    # x0 >= 700 = reference

    def parse(self, file_path: Path) -> ParsedStatement:
        stmt = ParsedStatement(bank_code=self.bank_code, bank_name=self.bank_name)

        with pdfplumber.open(file_path) as pdf:
            self._parse_header(pdf.pages[0], stmt)
            for page in pdf.pages:
                self._parse_transactions(page, stmt)

        return stmt

    @staticmethod
    def _group_lines(words: list) -> list[tuple[float, list]]:
        """Group words by approximate y-position, return [(top, [words])]."""
        lines: dict[int, list] = {}
        for w in words:
            key = round(w["top"])
            lines.setdefault(key, []).append(w)
        return [
            (top, sorted(lines[top], key=lambda x: x["x0"]))
            for top in sorted(lines)
        ]

    def _parse_header(self, page, stmt: ParsedStatement) -> None:
        text = page.extract_text() or ""

        m = re.search(
            r"Izvod\s+broj\s+(\d+)\s+za\s+promet\s+i\s+stanje\s+računa\s+(\d{18})\s+na\s+dan\s+(\d{2}\.\d{2}\.\d{4})",
            text,
        )
        if m:
            stmt.statement_number = m.group(1)
            stmt.account_number = m.group(2)
            stmt.statement_date = self.parse_date_dmy(m.group(3))

        m = re.search(r"Matični\s+broj\s+(\d+)\s+Naziv\s+(.+?)\s+Adresa\s+", text)
        if m:
            stmt.client_pib = m.group(1)
            stmt.client_name = self.clean_text(m.group(2)).strip('"')

        m = re.search(r"PIB\s+(\d+)", text)
        if m and not stmt.client_pib:
            stmt.client_pib = m.group(1)

        # Balances row
        m = re.search(
            r"Prethodno\s+stanje[^\n]*\n\s*"
            r"(\d[\d.]*,\d{2})\s+(\d[\d.]*,\d{2})\s+(\d[\d.]*,\d{2})\s+(\d[\d.]*,\d{2})",
            text,
        )
        if m:
            stmt.opening_balance = self.parse_amount_eu(m.group(1))
            stmt.total_debit = self.parse_amount_eu(m.group(2))
            stmt.total_credit = self.parse_amount_eu(m.group(3))
            stmt.closing_balance = self.parse_amount_eu(m.group(4))

    def _parse_transactions(self, page, stmt: ParsedStatement) -> None:
        words = page.extract_words(keep_blank_chars=False)
        line_groups = self._group_lines(words)

        # Find row lines: Rbr digit at x0 ≈ 62 followed by 10-digit ID at x0 ≈ 93
        row_indices = []
        for idx, (top, ws) in enumerate(line_groups):
            if (
                len(ws) >= 3
                and ws[0]["text"].isdigit()
                and len(ws[0]["text"]) <= 3
                and ws[0]["x0"] < 80
                and ws[1]["text"].isdigit()
                and len(ws[1]["text"]) >= 8
                and ws[1]["x0"] < 160
            ):
                row_indices.append(idx)

        for k, idx in enumerate(row_indices):
            top, ws = line_groups[idx]
            row_num = int(ws[0]["text"])

            transaction_id = ws[1]["text"]
            account_raw = ws[2]["text"] if len(ws) > 2 else ""

            payment_code = None
            value_date = None
            booking_date = None
            debit = None
            credit = None
            fee = None
            reference = None

            for w in ws[3:]:
                t = w["text"]
                x0 = w["x0"]

                # Šifra plaćanja: 3 digits at x ≈ 277
                if 270 <= x0 < 320 and re.fullmatch(r"\d{2,4}", t):
                    payment_code = t
                # Date: dd/mm/yyyy
                elif re.fullmatch(r"\d{2}/\d{2}/\d{4}", t):
                    if booking_date is None:
                        booking_date = self.parse_date_dmy_slash(t)
                    elif value_date is None:
                        value_date = self.parse_date_dmy_slash(t)
                # Amount: NN.NNN,NN
                elif re.fullmatch(r"\d[\d.]*,\d{2}", t):
                    amt = self.parse_amount_eu(t)
                    if x0 < self.X_ODLIV_MAX:
                        debit = amt
                    elif x0 < self.X_PRILIV_MAX:
                        credit = amt
                    elif x0 < self.X_PROVIZIJA_MAX:
                        fee = amt
                # Reference: e.g. "89-72-99-...", "28719-PPT"
                elif x0 >= self.X_PROVIZIJA_MAX:
                    reference = t if reference is None else f"{reference} {t}"

            if value_date is None:
                value_date = booking_date

            # Lines after row line until next row = counterparty + purpose
            # Each line may have left-column text (cp/purpose) and right-column reference
            next_idx = row_indices[k + 1] if k + 1 < len(row_indices) else len(line_groups)

            cp_lines: list[str] = []
            extra_refs: list[str] = []
            for j in range(idx + 1, next_idx):
                top_j, ws_j = line_groups[j]
                if ws_j and ws_j[0]["text"] == "UKUPNO":
                    break
                left = " ".join(w["text"] for w in ws_j if w["x0"] < self.X_PROVIZIJA_MAX).strip()
                right = " ".join(w["text"] for w in ws_j if w["x0"] >= self.X_PROVIZIJA_MAX).strip()
                if left:
                    cp_lines.append(left)
                if right:
                    extra_refs.append(right)

            counterparty = self.clean_text(cp_lines[0]) if cp_lines else None
            purpose = self.clean_text(" ".join(cp_lines[1:])) if len(cp_lines) > 1 else None

            txn = ParsedTransaction(
                row_number=row_num,
                value_date=value_date,
                booking_date=booking_date,
                debit=debit,
                credit=credit,
                counterparty=counterparty,
                counterparty_account=account_raw if re.fullmatch(r"\d{15,18}", account_raw) else None,
                payment_code=payment_code,
                purpose=purpose,
                fee=fee,
            )
            if reference:
                if debit is not None:
                    txn.reference_debit = reference
                else:
                    txn.reference_credit = reference
            if extra_refs:
                combined = " ".join(extra_refs)
                if debit is not None and not txn.reference_debit:
                    txn.reference_debit = combined
                elif credit is not None and not txn.reference_credit:
                    txn.reference_credit = combined

            stmt.transactions.append(txn)
