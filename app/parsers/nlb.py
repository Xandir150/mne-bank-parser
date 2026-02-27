import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction


# CID -> character maps for NLB Banka PDFs with Identity-H encoded fonts
_BOLD_CID_MAP = {
    4: 'I', 5: 'Z', 6: 'V', 7: 'O', 8: 'D', 9: 'B', 10: 'R', 11: '.',
    12: '1', 13: 'A', 14: 'P', 15: 'M', 16: 'J', 17: 'E', 18: 'N', 19: 'U',
    20: 'S', 21: 'T', 22: 'C', 23: '6', 24: '0', 25: '2', 26: '5', 27: '3',
    28: '-', 29: '9', 30: '7', 31: 'o', 32: 'k', 33: 'r', 34: 'i', 35: 'ć',
    36: 'e', 37: 'u', 38: 'p', 39: 'n',
}

_REGULAR_CID_MAP = {
    4: '(', 5: 'N', 6: 'a', 7: 'z', 8: 'i', 9: 'v', 10: 'l', 11: 's',
    12: 'n', 13: 'k', 14: 'r', 15: 'č', 16: 'u', 17: ')', 18: 'B', 19: 'o',
    20: 'j', 21: 'P', 22: 'e', 23: 't', 24: 'h', 25: 'd', 26: 'D', 27: 'm',
    28: 'g', 29: 'p', 30: 'b', 31: '0', 32: '3', 33: '4', 34: '9', 35: '5',
    36: '7', 37: 'ž', 38: 'I', 39: 'š', 40: 'ć', 41: '.', 42: 'R', 43: 'T',
    44: '1', 45: '6', 46: 'c', 47: '-', 48: 'Š', 49: 'f', 50: 'S', 51: 'L',
    52: 'A', 53: ',', 54: 'C', 55: 'E', 56: 'K', 57: '2', 58: ':', 59: 'U',
}

# Minimum x-gap between characters to insert a space
_SPACE_THRESHOLD = 2.0


@register_parser
class NLBParser(BankParser):
    """Parser for NLB Banka (530) PDF statements.

    NLB PDFs use custom Identity-H font encoding that requires CID-to-character
    mapping. The parser decodes characters, groups them into words, then parses
    the structured table layout.

    Columns: nal.br, counterparty+account, origin+date, debit, credit,
    payment code, purpose, references, reclamation.
    Period decimal (109.66). Date DD.MM.YYYY.
    """

    bank_code = "530"
    bank_name = "NLB Banka"

    def parse(self, file_path: Path) -> ParsedStatement:
        stmt = ParsedStatement(
            bank_code=self.bank_code,
            bank_name=self.bank_name,
        )

        with pdfplumber.open(file_path) as pdf:
            for page_num, page in enumerate(pdf.pages):
                decoded_lines = self._decode_page(page)
                if page_num == 0:
                    self._parse_header(decoded_lines, page, stmt)
                self._parse_transactions(decoded_lines, page, stmt)

        return stmt

    def _decode_char(self, char_obj: dict) -> tuple[Optional[str], bool]:
        """Decode a single character from its CID representation.
        Returns (character, is_bold).
        """
        text = char_obj.get("text", "")
        fontname = char_obj.get("fontname", "")
        is_bold = "Bold" in fontname or "bold" in fontname

        m = re.match(r"\(cid:(\d+)\)", text)
        if not m:
            return text, is_bold

        cid = int(m.group(1))
        if is_bold:
            return _BOLD_CID_MAP.get(cid, ""), True
        else:
            return _REGULAR_CID_MAP.get(cid, ""), False

    def _decode_page(self, page) -> list[tuple[float, list[tuple[float, str]]]]:
        """Decode all characters on a page and group into lines of words.

        Returns list of (y_position, [(x_position, text), ...]) tuples.
        """
        chars = page.chars
        if not chars:
            return []

        # Decode each character with font info
        decoded = []
        for c in chars:
            result = self._decode_char(c)
            ch, is_bold = result[0], result[1]
            if ch and isinstance(ch, str):
                decoded.append((c["x0"], c["top"], str(ch), is_bold, c.get("width", 5)))

        # Group by y-position (within 4pt tolerance)
        from collections import defaultdict
        y_groups = defaultdict(list)
        for x, y, ch, is_bold, w in decoded:
            y_key = round(y / 4) * 4
            y_groups[y_key].append((x, ch, is_bold, w))

        # For each line, combine characters into words
        # Bold font has wider spacing, so use the actual character width
        result = []
        for y_key in sorted(y_groups.keys()):
            chars_on_line = sorted(y_groups[y_key], key=lambda p: p[0])
            words = []
            current_word = ""
            current_x = None
            current_end_x = None
            word_start_x = None

            for item in chars_on_line:
                x, ch, is_bold, char_w = item[0], str(item[1]), item[2], item[3]
                if current_end_x is not None and (x - current_end_x) > _SPACE_THRESHOLD:
                    if current_word:
                        words.append((word_start_x, current_word))
                    current_word = ch
                    word_start_x = x
                else:
                    if not current_word:
                        word_start_x = x
                    current_word += ch
                current_end_x = x + max(char_w, 3)

            if current_word:
                words.append((word_start_x, current_word))

            result.append((y_key, words))

        return result

    def _line_text(self, words: list[tuple[float, str]]) -> str:
        """Join words from a line into a single string."""
        return " ".join(w[1] for w in words)

    def _parse_header(self, lines: list, page, stmt: ParsedStatement) -> None:
        for y, words in lines:
            text = self._line_text(words)

            # Title: "IZVOD BR. 1"
            m = re.search(r"IZVOD\s*BR\.\s*(\d+)", text)
            if m:
                stmt.statement_number = m.group(1)

            # Date: "ZA PROMJENU SREDSTAVA NA RACUNU DANA DD.MM.YYYY"
            m = re.search(r"DANA\s+(\d{2}\.\d{2}\.\d{4})", text)
            if m:
                stmt.statement_date = self.parse_date_dmy(m.group(1))

            # Client name: "NURBAN" (bold, left side, early y)
            # Look for the client name line (before "STANJE" and before account)
            if y < 120:
                for x, w in words:
                    if x < 200 and re.match(r"^[A-Z][A-Z\s]+$", w) and len(w) > 2:
                        if w not in ("IZVOD", "ZA", "STANJE", "NLB"):
                            stmt.client_name = self.clean_text(w)

            # Account: "530-0000000030153-55" (bold, right side)
            for x, w in words:
                if re.match(r"530-\d{13}-\d{2}", w):
                    stmt.account_number = w

            # PIB: "poreski broj 03349357"
            m = re.search(r"poreski\s*broj\s*(\d+)", text, re.IGNORECASE)
            if m:
                stmt.client_pib = m.group(1)

            # Balances: "109.66  2.00  107.66  1  0"
            # This is the summary row with period-decimal amounts
            amounts = re.findall(r"(\d+\.\d{2})", text)
            if len(amounts) >= 3 and y > 180 and y < 260:
                stmt.opening_balance = self.parse_amount_us(amounts[0])
                stmt.total_debit = self.parse_amount_us(amounts[1])
                if len(amounts) >= 4:
                    stmt.closing_balance = self.parse_amount_us(amounts[2])

    def _parse_transactions(self, lines: list, page, stmt: ParsedStatement) -> None:
        """Parse transactions from decoded lines.

        NLB transaction blocks span multiple lines:
        - Line with counterparty name (x~50-200)
        - Line with origin info + debit amount (x~216 for origin, x~341 for amount)
        - Line with row# + continuation + code + purpose + refs + reclam
        - Line with date + Naknada (fee)
        - Line with account number

        The row-number line is the key indicator of a transaction.
        """
        # Find PROMJENE section
        promjene_idx = None
        for i, (y, words) in enumerate(lines):
            text = self._line_text(words)
            combined = text.upper().replace(" ", "")
            if "PROMJENE" in combined:
                promjene_idx = i
                break

        if promjene_idx is None:
            return

        # Skip header rows after PROMJENE
        data_lines = lines[promjene_idx + 1:]

        # Find all row-number lines (transaction key lines)
        txn_key_indices = []
        ukupno_idx = len(data_lines)
        for i, (y, words) in enumerate(data_lines):
            text = self._line_text(words)
            if not words:
                continue
            first_word = words[0][1]
            # Row number: single digit at x < 40
            if first_word.isdigit() and words[0][0] < 45 and int(first_word) > 0:
                txn_key_indices.append(i)
            if "Ukupno" in text:
                ukupno_idx = i
                # Extract total from "Ukupno EURA 2.00"
                m = re.search(r"(\d+\.\d{2})", text)
                if m:
                    stmt.total_debit = self.parse_amount_us(m.group(1))
                break

        # For each transaction, collect its block of lines
        for t_idx, key_i in enumerate(txn_key_indices):
            # Block extends from 2 lines before key to next key (or Ukupno)
            block_start = max(0, key_i - 2)
            if t_idx + 1 < len(txn_key_indices):
                block_end = txn_key_indices[t_idx + 1] - 2  # leave room for next counterparty
            else:
                block_end = ukupno_idx

            # Ensure we don't overlap with previous transaction's key line
            if t_idx > 0:
                prev_key = txn_key_indices[t_idx - 1]
                block_start = max(block_start, prev_key + 1)

            block = data_lines[block_start:block_end]
            txn = self._parse_txn_block(block, data_lines[key_i])
            if txn:
                stmt.transactions.append(txn)

    def _parse_txn_block(self, block: list, key_line: tuple) -> Optional[ParsedTransaction]:
        """Parse a single transaction from its block of lines."""
        key_y, key_words = key_line
        if not key_words:
            return None

        first_word = key_words[0][1]
        if not first_word.isdigit():
            return None

        txn = ParsedTransaction(row_number=int(first_word))

        # Parse the key line (row# + code + purpose + refs + reclam)
        # x-ranges based on NLB table column positions:
        # <210: counterparty cont.  210-340: origin  340-413: amounts
        # 413-445: payment code  445-650: purpose  650-740: refs  740+: reclam
        for x, w in key_words[1:]:
            if x < 210:
                pass  # counterparty continuation (like "46,")
            elif x < 340:
                pass  # origin continuation
            elif x < 413:
                if w not in ("-", "-:-"):
                    pass  # amount
            elif x < 445:
                if not txn.payment_code:
                    txn.payment_code = w
            elif x < 650:
                if txn.purpose:
                    txn.purpose += " " + w
                else:
                    txn.purpose = w
            elif x < 740:
                txn.reference_debit = w
            else:
                if txn.reclamation_data:
                    txn.reclamation_data += w
                else:
                    txn.reclamation_data = w

        # Parse all lines in the block for additional data
        for y, words in block:
            if (y, words) == key_line:
                continue
            text = self._line_text(words)

            # Skip table header lines
            lower = text.lower()
            if any(h in lower for h in ("nal.", "naziv", "sjedišt", "šifra",
                                         "datum", "knjižen", "zaduženje", "odobrenje",
                                         "iznos", "svrha", "poziv", "podaci",
                                         "reklamacij")):
                continue

            # Naknada (fee) line - check first to avoid misidentifying fee as debit
            is_fee_line = "Naknada" in text
            if is_fee_line:
                m = re.search(r"Naknada\s*([\d.]+)", text)
                if m:
                    txn.fee = self.parse_amount_us(m.group(1))

            # Counterparty name line (mostly text at x < 210)
            name_words = [w for x, w in words if x < 210]
            if name_words and not txn.counterparty:
                txn.counterparty = self.clean_text(" ".join(name_words))
            elif name_words and txn.counterparty:
                extra = self.clean_text(" ".join(name_words))
                if extra and not re.match(r"^\d{3}-", extra):
                    txn.counterparty += " " + extra

            # Debit amount (typically at x~341) - skip Naknada lines
            if not is_fee_line:
                for x, w in words:
                    if 300 < x < 380 and re.match(r"^\d+\.\d{2}$", w):
                        txn.debit = self.parse_amount_us(w)
                        if txn.debit == Decimal("0"):
                            txn.debit = None

            # Date line: "16.01.2026"
            for x, w in words:
                m = re.match(r"(\d{2}\.\d{2}\.\d{4})", w)
                if m:
                    txn.booking_date = self.parse_date_dmy(m.group(1))
                    txn.value_date = txn.booking_date

            # Account number
            for x, w in words:
                if re.match(r"\d{3}-\d{13}-\d{2}$", w):
                    txn.counterparty_account = w

        if txn.purpose:
            txn.purpose = self.clean_text(txn.purpose)

        if txn.debit is not None or txn.credit is not None:
            return txn
        return None
