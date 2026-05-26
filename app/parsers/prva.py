import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


@register_parser
class PrvaParser(BankParser):
    """Parser for Prva Banka CG (535) PDF statements.

    Layout: RB-numbered rows with Naknada (fee) in the debit column.
    Columns: RB, counterparty+account, origin+date, debit, naknada, credit,
    code, purpose, references, reclamation.
    Comma decimal (2.163,40) for main amounts.
    Period decimal (0.44) for fees.
    Date YYYY.MM.DD.
    """

    bank_code = "535"
    bank_name = "Prva Banka CG"

    # Regex for a transaction start line: "RB NAME... AMOUNT AMOUNT CODE PURPOSE..."
    # RB is a small number (1-999), followed by at least one alpha char
    _TXN_START = re.compile(
        r"^(\d{1,3})\s+([A-Za-z/\",].+?)\s+([\d.]+,\d{2})\s+([\d.]+,\d{2})\s+(\d{3})\s+(.*)"
    )

    def parse(self, file_path: Path) -> ParsedStatement:
        stmt = ParsedStatement(
            bank_code=self.bank_code,
            bank_name=self.bank_name,
        )

        with pdfplumber.open(file_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"

            # Detect new "IZVOD BROJ N / STANJE I PROMJENE SREDSTAVA" format
            # (font with custom CID encoding, column-based layout)
            is_new_format = bool(
                re.search(r"IZVOD\s+BROJ\s+\d+", full_text)
                and re.search(r"STANJE\s+I\s+PROMJENE\s+SREDSTAVA\s+NA\s+DAN", full_text)
            )
            if is_new_format:
                self._parse_new_header(full_text, stmt)
                for page in pdf.pages:
                    self._parse_new_transactions(page, stmt)
                if stmt.transactions:
                    return stmt
                # fall through to legacy parsing if no transactions extracted

            self._parse_header(full_text, stmt)
            self._parse_transactions(full_text, stmt)

        return stmt

    # ----------------------------------------------------------------
    # New format ("IZVOD BROJ N / STANJE I PROMJENE SREDSTAVA NA DAN")
    # Column-based layout, custom font CID encoding.
    # ----------------------------------------------------------------
    # Column x-boundaries (chars are right-aligned in numeric cols)
    _COL_RB = (35, 50)
    _COL_LEFT = (50, 190)        # counterparty / city / account
    _COL_MID = (190, 320)        # origin / date_knjizenja / date_prijema
    _COL_NAKNADA = (320, 360)    # "Naknada:  X.XX" label/value combined
    _COL_DEBIT = (340, 405)      # duguje (right-aligned, ends ~385)
    _COL_CREDIT = (405, 470)     # potražuje (ends ~460)
    _COL_SIFRA = (465, 495)      # 3-digit code
    _COL_PURPOSE = (495, 605)
    _COL_REF_RIGHT = (605, 680)  # poziv na broj (odobrenja) / extra ref
    _COL_RECLAMATION = (680, 760)

    @staticmethod
    def _strip_cids(text: str) -> str:
        return re.sub(r"\(cid:\d+\)", "", text)

    def _parse_new_header(self, text: str, stmt: ParsedStatement) -> None:
        text = self._strip_cids(text)

        m = re.search(r"IZVOD\s+BROJ\s+(\d+)", text)
        if m:
            stmt.statement_number = m.group(1)

        m = re.search(r"STANJE\s+I\s+PROMJENE\s+SREDSTAVA\s+NA\s+DAN\s+(\d{2}\.\d{2}\.\d{4})", text)
        if m:
            stmt.statement_date = self.parse_date_dmy(m.group(1))

        # Client name + account on one line: "Civitas Solis doo Budva 535-0000000021278-71"
        m = re.search(r"^\s*([A-Z][^\n]*?)\s+(535-\d+-\d{2})\s*$", text, re.MULTILINE)
        if m:
            stmt.client_name = self.clean_text(m.group(1))
            stmt.account_number = self.normalize_account(m.group(2))

        m = re.search(r"Poreski\s+broj:\s*(\d+)", text)
        if m:
            stmt.client_pib = m.group(1)

        # Balance row: opening, debit, credit, closing, num_debit, num_credit
        # e.g. "67.99 7.00 0.00 60.99 2 0" or "1,060.99 286.24 0.00 774.75 6 0"
        m = re.search(
            r"prethodno\s+stanje\s+duguje\s+potražuje\s+novo\s+stanje\s+duguje\s+potražuje[\s\S]*?"
            r"\n\s*([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+(\d+)\s+(\d+)\b",
            text,
        )
        if m:
            stmt.opening_balance = self.parse_amount_us(m.group(1))
            stmt.total_debit = self.parse_amount_us(m.group(2))
            stmt.total_credit = self.parse_amount_us(m.group(3))
            stmt.closing_balance = self.parse_amount_us(m.group(4))

    @staticmethod
    def _group_chars_into_segments(chars: list) -> list[tuple[float, str]]:
        """Group consecutive chars (close x positions) into (x_start, text) segments."""
        if not chars:
            return []
        chars = sorted(chars, key=lambda c: c["x0"])
        segments: list[tuple[float, str]] = []
        cur_text = ""
        cur_x = chars[0]["x0"]
        prev_x = None
        for c in chars:
            if prev_x is not None and (c["x0"] - prev_x) > 3:
                if cur_text.strip():
                    segments.append((cur_x, cur_text))
                cur_text = c["text"]
                cur_x = c["x0"]
            else:
                if not cur_text:
                    cur_x = c["x0"]
                cur_text += c["text"]
            prev_x = c["x1"]
        if cur_text.strip():
            segments.append((cur_x, cur_text))
        # Strip CIDs from segments
        cleaned: list[tuple[float, str]] = []
        for x, t in segments:
            t2 = re.sub(r"\(cid:\d+\)", "", t).strip()
            if t2:
                cleaned.append((x, t2))
        return cleaned

    def _parse_new_transactions(self, page, stmt: ParsedStatement) -> None:
        # Group chars by y (rounded to int)
        chars_by_y: dict[int, list] = {}
        for c in page.chars:
            key = round(c["top"])
            chars_by_y.setdefault(key, []).append(c)

        # Build per-row segment list
        rows: list[tuple[int, list[tuple[float, str]]]] = []
        for y in sorted(chars_by_y):
            segs = self._group_chars_into_segments(chars_by_y[y])
            if segs:
                rows.append((y, segs))

        # Find rb. markers: a digit+dot at col_rb position
        rb_re = re.compile(r"^(\d{1,3})\.$")
        rb_indices: list[tuple[int, int, int]] = []  # (row_idx, y, rb_num)
        for idx, (y, segs) in enumerate(rows):
            for x, t in segs:
                if self._COL_RB[0] <= x < self._COL_RB[1]:
                    m = rb_re.match(t)
                    if m:
                        rb_indices.append((idx, y, int(m.group(1))))
                        break

        for k, (row_idx, rb_y, rb_num) in enumerate(rb_indices):
            # Block y-span: from previous block end (or header end) to next block's rb_y
            prev_y = rb_indices[k - 1][1] if k > 0 else 0
            next_y = rb_indices[k + 1][1] if k + 1 < len(rb_indices) else 10_000

            block_top = (prev_y + rb_y) // 2 if k > 0 else rb_y - 12
            block_bottom = (rb_y + next_y) // 2 if k + 1 < len(rb_indices) else rb_y + 18

            cp_parts: list[str] = []
            account: Optional[str] = None
            origin_parts: list[str] = []
            booking_date = None
            value_date = None
            debit = None
            credit = None
            fee = None
            payment_code = None
            purpose_parts: list[str] = []
            extra_ref_parts: list[str] = []
            reclamation: Optional[str] = None

            for y, segs in rows:
                if not (block_top <= y <= block_bottom):
                    continue
                for x, t in segs:
                    # RB column — skip
                    if self._COL_RB[0] <= x < self._COL_RB[1]:
                        continue
                    # LEFT column: counterparty / city / account
                    if self._COL_LEFT[0] <= x < self._COL_LEFT[1]:
                        if re.fullmatch(r"\d{3}-\d+-\d{2}", t):
                            account = self.normalize_account(t)
                        elif t == "-":
                            pass
                        else:
                            cp_parts.append(t)
                        continue
                    # MID column: origin / date_knj / date_prijema
                    if self._COL_MID[0] <= x < self._COL_MID[1]:
                        md = re.match(r"^(\d{1,2}\.\d{1,2}\.\d{4})$", t)
                        if md:
                            d = self.parse_date_dmy(md.group(1))
                            if d and booking_date is None:
                                booking_date = d
                            elif d:
                                value_date = d
                        else:
                            origin_parts.append(t)
                        continue
                    # NAKNADA combined "Naknada: X.XX"
                    if self._COL_NAKNADA[0] <= x < self._COL_NAKNADA[1]:
                        mn = re.search(r"Naknada:\s*([\d,]+\.\d{2})", t)
                        if mn:
                            fee = self.parse_amount_us(mn.group(1))
                            continue
                        # else maybe a debit value sliding into this range
                    # DEBIT
                    if self._COL_DEBIT[0] <= x < self._COL_DEBIT[1]:
                        if re.fullmatch(r"[\d,]+\.\d{2}", t):
                            debit = self.parse_amount_us(t)
                            continue
                    # CREDIT
                    if self._COL_CREDIT[0] <= x < self._COL_CREDIT[1]:
                        if re.fullmatch(r"[\d,]+\.\d{2}", t):
                            credit = self.parse_amount_us(t)
                            continue
                    # SIFRA
                    if self._COL_SIFRA[0] <= x < self._COL_SIFRA[1]:
                        if re.fullmatch(r"\d{2,4}", t):
                            payment_code = t
                            continue
                    # PURPOSE
                    if self._COL_PURPOSE[0] <= x < self._COL_PURPOSE[1]:
                        purpose_parts.append(t)
                        continue
                    # POZIV NA BROJ (odobrenja) / extra ref
                    if self._COL_REF_RIGHT[0] <= x < self._COL_REF_RIGHT[1]:
                        extra_ref_parts.append(t)
                        continue
                    # RECLAMATION
                    if self._COL_RECLAMATION[0] <= x < self._COL_RECLAMATION[1]:
                        if re.fullmatch(r"\d{8,}", t) and reclamation is None:
                            reclamation = t

            # Only zero-zero pairs are not real txns; skip if both 0 and no fee
            if debit is None and credit is None:
                continue
            if debit is not None and debit == Decimal("0"):
                debit = None
            if credit is not None and credit == Decimal("0"):
                credit = None

            origin = self.clean_text(" ".join(origin_parts)) if origin_parts else None
            counterparty = self.clean_text(" ".join(cp_parts)) if cp_parts else None
            purpose = self.clean_text(" ".join(purpose_parts)) if purpose_parts else None

            txn = ParsedTransaction(
                row_number=rb_num,
                value_date=value_date or booking_date,
                booking_date=booking_date,
                debit=debit,
                credit=credit,
                counterparty=counterparty,
                counterparty_account=account,
                counterparty_bank=origin,
                payment_code=payment_code,
                purpose=purpose,
                reclamation_data=reclamation,
                fee=fee,
            )
            if extra_ref_parts:
                ref = self.clean_text(" ".join(extra_ref_parts))
                if credit is not None:
                    txn.reference_credit = ref
                else:
                    txn.reference_debit = ref
            stmt.transactions.append(txn)

    # ----------------------------------------------------------------
    # Legacy format
    # ----------------------------------------------------------------
    def _parse_header(self, text: str, stmt: ParsedStatement) -> None:
        # Client name
        m = re.search(r"Naziv:\s*(.+?)(?:\s+Izvod|\s*$)", text, re.MULTILINE)
        if m:
            stmt.client_name = self.clean_text(m.group(1))

        # PIB (client's, not bank's) - get the second PIB occurrence
        pib_matches = re.findall(r"PIB:\s*(\d+)", text)
        if len(pib_matches) >= 2:
            stmt.client_pib = pib_matches[1]
        elif pib_matches:
            stmt.client_pib = pib_matches[0]

        # Account - get the client account (second Račun: line), not the bank's
        acct_matches = re.findall(r"Ra[čc]un:\s*(535-[\d-]+)", text)
        if len(acct_matches) >= 2:
            stmt.account_number = acct_matches[1]
        elif acct_matches:
            stmt.account_number = acct_matches[0]

        # Statement number
        m = re.search(r"SREDSTAVA\s+BROJ\s+(\d+)", text)
        if m:
            stmt.statement_number = m.group(1)

        # Statement date
        m = re.search(r"Datum\s+izvoda:\s*(\d{2}\.\d{2}\.\d{4})", text)
        if m:
            stmt.statement_date = self.parse_date_dmy(m.group(1))

        # Balance summary: "2.163,40 1.083,99 0,00 1.079,41 6 / 0 ..."
        m = re.search(
            r"([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+\d+\s*/\s*\d+",
            text,
        )
        if m:
            stmt.opening_balance = self.parse_amount_eu(m.group(1))
            stmt.total_debit = self.parse_amount_eu(m.group(2))
            stmt.total_credit = self.parse_amount_eu(m.group(3))
            stmt.closing_balance = self.parse_amount_eu(m.group(4))

    def _parse_transactions(self, text: str, stmt: ParsedStatement) -> None:
        lines = text.split("\n")

        # Find all transaction start indices
        txn_starts = []
        for i, line in enumerate(lines):
            m = self._TXN_START.match(line.strip())
            if m:
                txn_starts.append(i)

        # Also find the UKUPNO line to bound the last transaction
        ukupno_idx = len(lines)
        for i, line in enumerate(lines):
            if line.strip().startswith("UKUPNO"):
                ukupno_idx = i
                break

        # Parse each transaction block
        for idx, start_i in enumerate(txn_starts):
            end_i = txn_starts[idx + 1] if idx + 1 < len(txn_starts) else ukupno_idx
            block = [lines[j].strip() for j in range(start_i, end_i) if lines[j].strip()]
            txn = self._parse_block(block)
            if txn:
                stmt.transactions.append(txn)

    def _parse_block(self, block: list[str]) -> Optional[ParsedTransaction]:
        if not block:
            return None

        first = block[0]
        m = self._TXN_START.match(first)
        if not m:
            return None

        rb = int(m.group(1))
        name_and_origin = m.group(2)
        debit_str = m.group(3)
        credit_str = m.group(4)
        payment_code = m.group(5)
        purpose_and_rest = m.group(6)

        txn = ParsedTransaction(row_number=rb)
        txn.payment_code = payment_code

        # Parse debit/credit
        txn.debit = self.parse_amount_eu(debit_str)
        txn.credit = self.parse_amount_eu(credit_str)
        if txn.debit is not None and txn.debit == Decimal("0"):
            txn.debit = None
        if txn.credit is not None and txn.credit == Decimal("0"):
            txn.credit = None

        # Parse purpose + references from the tail
        # Pattern: "purpose_text ( ) RECLAMATION" or "purpose_text (MODEL) REF"
        self._parse_purpose_tail(purpose_and_rest, txn)

        # Parse counterparty name from name_and_origin
        # It's "NAME... Filijala/origin..."
        m_name = re.match(r"(.+?)\s+(?:Filijala\b|0\d{3}\b)", name_and_origin)
        if m_name:
            txn.counterparty = self.clean_text(m_name.group(1))
        else:
            txn.counterparty = self.clean_text(name_and_origin)

        # Parse continuation lines
        for line in block[1:]:
            # Account + date: "820-30000-74 2026.02.24"
            m = re.match(r"(\d{3}-[\d-]+)\s+(\d{4}\.\d{2}\.\d{2})", line)
            if m:
                txn.counterparty_account = m.group(1)
                txn.booking_date = self.parse_date_ymd(m.group(2))
                txn.value_date = txn.booking_date
                continue

            # Just date: "2026.02.24"
            m = re.match(r"^(\d{4}\.\d{2}\.\d{2})\s*$", line)
            if m:
                txn.booking_date = self.parse_date_ymd(m.group(1))
                txn.value_date = txn.booking_date
                continue

            # Account with origin and fee: "530-54171-72 0431 0.34 ( ) 03244822"
            m = re.match(r"(\d{3}-[\d-]+)\s+\d{4}\s+([\d.]+)\s+\(([^)]*)\)\s*(.*)", line)
            if m:
                txn.counterparty_account = m.group(1)
                txn.fee = self.parse_amount_us(m.group(2))
                ref_model = m.group(3).strip()
                ref_val = m.group(4).strip()
                if ref_model or ref_val:
                    txn.reference_credit = self.clean_text(f"({ref_model}) {ref_val}")
                continue

            # Fee line without account: "0431 0.44 (18) 03486575-302"
            m = re.match(r"(\d{4})\s+([\d.]+)\s+\(([^)]*)\)\s*(.*)", line)
            if m:
                txn.fee = self.parse_amount_us(m.group(2))
                ref_model = m.group(3).strip()
                ref_val = m.group(4).strip()
                if ref_model or ref_val:
                    txn.reference_credit = self.clean_text(f"({ref_model}) {ref_val}")
                continue

            # Continuation of counterparty name (lines with alpha text, before account)
            if not txn.counterparty_account and re.match(r"^[A-Za-z/,]", line):
                # This may be a continuation of the name
                extra_name = line.split("stari ")[0].split("Filijala")[0].strip()
                extra_name = re.sub(r"\(\d+\)$", "", extra_name).strip()
                if extra_name:
                    txn.counterparty = (txn.counterparty or "") + " " + self.clean_text(extra_name)
                    txn.counterparty = self.clean_text(txn.counterparty)

        if txn.debit is not None or txn.credit is not None:
            return txn
        return None

    def _parse_purpose_tail(self, tail: str, txn: ParsedTransaction) -> None:
        """Parse purpose text and references from the tail of the first line."""
        # Pattern: "purpose_text ( ) RECLAMATION_NUMBER"
        # Or: "purpose_text (MODEL) REF_NUMBER"
        m = re.search(r"\(\s*\)\s+(\d{11,})", tail)
        if m:
            txn.purpose = self.clean_text(tail[: m.start()])
            txn.reclamation_data = m.group(1)
            return

        m = re.search(r"\(([^)]*)\)\s+(\d{11,})", tail)
        if m:
            txn.purpose = self.clean_text(tail[: m.start()])
            txn.reference_debit = f"({m.group(1)})"
            txn.reclamation_data = m.group(2)
            return

        txn.purpose = self.clean_text(tail) if tail else None
