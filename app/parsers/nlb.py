import logging
import re
from collections import defaultdict
from decimal import Decimal
from io import BytesIO
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction

logger = logging.getLogger(__name__)

# Minimum x-gap between characters to insert a space
_SPACE_THRESHOLD = 2.0

# ---------------------------------------------------------------------------
# Reference font paths (Arial or Liberation Sans -- metrically identical)
# ---------------------------------------------------------------------------

_ARIAL_REGULAR_PATHS = [
    # macOS
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial.ttf",
    # Linux / Docker (fonts-liberation package)
    "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
    # Linux (msttcorefonts)
    "/usr/share/fonts/truetype/msttcorefonts/Arial.ttf",
]

_ARIAL_BOLD_PATHS = [
    # macOS
    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    "/Library/Fonts/Arial Bold.ttf",
    # Linux / Docker (fonts-liberation package)
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
    # Linux (msttcorefonts)
    "/usr/share/fonts/truetype/msttcorefonts/Arial_Bold.ttf",
]


def _find_font(paths: list[str]) -> Optional[str]:
    """Return the first existing font path, or None."""
    for p in paths:
        if Path(p).is_file():
            return p
    return None


def _get_arial_hashes(bold: bool = False) -> dict[str, str]:
    """Get hash -> character mapping for Arial glyph outlines.

    First tries to load pre-built hashes from bundled JSON (works everywhere).
    Falls back to runtime extraction from system fonts if JSON not found.
    """
    # Try bundled pre-built hashes first
    json_path = Path(__file__).parent / "arial_hashes.json"
    if json_path.is_file():
        import json
        with open(json_path, encoding="utf-8") as f:
            data = json.load(f)
        key = "bold" if bold else "regular"
        return data.get(key, {})

    # Fallback: runtime extraction from system font
    from fontTools.pens.hashPointPen import HashPointPen
    from fontTools.ttLib import TTFont

    paths = _ARIAL_BOLD_PATHS if bold else _ARIAL_REGULAR_PATHS
    font_path = _find_font(paths)
    if font_path is None:
        style = "bold" if bold else "regular"
        logger.warning(
            "No reference %s font found and no arial_hashes.json. Tried: %s", style, paths
        )
        return {}

    ttf = TTFont(font_path)
    gs = ttf.getGlyphSet()

    cmap = ttf.getBestCmap()
    name_to_char: dict[str, str] = {}
    if cmap:
        for codepoint, glyph_name in sorted(cmap.items()):
            ch = chr(codepoint)
            existing = name_to_char.get(glyph_name)
            if existing is None:
                name_to_char[glyph_name] = ch
            elif ord(ch) < 128 and ch.isprintable() and (
                ord(existing) >= 128 or not existing.isprintable()
            ):
                name_to_char[glyph_name] = ch

    _SKIP_CHARS = frozenset(["\u200b"])
    result: dict[str, str] = {}
    for glyph_name in ttf.getGlyphOrder():
        ch = name_to_char.get(glyph_name)
        if ch is None or ch in _SKIP_CHARS:
            continue
        try:
            pen = HashPointPen(glyphSet=gs)
            gs[glyph_name].drawPoints(pen)
            h = pen.hash
            existing = result.get(h)
            if existing is None:
                result[h] = ch
            elif ord(ch) < 128 and ch.isprintable() and (
                ord(existing) >= 128 or not existing.isprintable()
            ):
                result[h] = ch
        except Exception:
            pass

    ttf.close()
    return result


def _build_cid_maps_from_font(pdf_path: Path) -> tuple[dict[int, str], dict[int, str]]:
    """Extract CID -> character maps by matching embedded font glyph outlines against Arial.

    Returns (bold_map, regular_map).
    """
    from pdfminer.pdfparser import PDFParser
    from pdfminer.pdfdocument import PDFDocument
    from pdfminer.pdfpage import PDFPage
    from pdfminer.pdftypes import resolve1
    from fontTools.pens.hashPointPen import HashPointPen
    from fontTools.ttLib import TTFont

    # Build reference hashes from system Arial / Liberation Sans
    ref_regular = _get_arial_hashes(bold=False)
    ref_bold = _get_arial_hashes(bold=True)

    if not ref_regular and not ref_bold:
        logger.error(
            "No reference fonts available. Install fonts-liberation package "
            "(Linux/Docker) or ensure Arial is available (macOS)."
        )

    bold_map: dict[int, str] = {}
    regular_map: dict[int, str] = {}

    with open(pdf_path, "rb") as f:
        parser = PDFParser(f)
        doc = PDFDocument(parser)
        for page in PDFPage.create_pages(doc):
            resources = resolve1(page.resources)
            fonts = resolve1(resources.get("Font", {}))
            for font_name, font_ref in fonts.items():
                font = resolve1(font_ref)
                encoding = str(font.get("Encoding", ""))
                if "Identity-H" not in encoding:
                    continue
                base_font = str(font.get("BaseFont", ""))
                is_bold = "Bold" in base_font
                ref = ref_bold if is_bold else ref_regular
                target_map = bold_map if is_bold else regular_map

                descendant_fonts = font.get("DescendantFonts", [])
                if descendant_fonts is None:
                    continue
                for d in resolve1(descendant_fonts):
                    desc = resolve1(d)
                    fd = resolve1(desc.get("FontDescriptor", {}))
                    if "FontFile2" not in fd:
                        continue
                    data = resolve1(fd["FontFile2"]).get_data()
                    ttf = TTFont(BytesIO(data))
                    gs = ttf.getGlyphSet()
                    for name in ttf.getGlyphOrder():
                        if name == ".notdef":
                            continue
                        # Glyph names in embedded NLB fonts are "glyphXXXXX"
                        m = re.match(r"glyph(\d+)", name)
                        if not m:
                            continue
                        gid = int(m.group(1))
                        try:
                            pen = HashPointPen(glyphSet=gs)
                            gs[name].drawPoints(pen)
                            ch = ref.get(pen.hash)
                            if ch:
                                target_map[gid] = ch
                        except Exception:
                            pass
                    ttf.close()
            break  # first page fonts are sufficient

    logger.info(
        "CID maps built: bold=%d entries, regular=%d entries",
        len(bold_map), len(regular_map),
    )
    return bold_map, regular_map


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------


@register_parser
class NLBParser(BankParser):
    """Parser for NLB Banka (530) PDF statements.

    NLB PDFs use custom Identity-H font encoding where text is encoded as
    (cid:XX).  The embedded TrueType fonts have the same glyph outlines as
    standard Arial, so we match outlines via HashPointPen against a reference
    Arial (or Liberation Sans) font to build CID -> character maps.

    Columns: nal.br, counterparty+account, origin+date, debit, credit,
    payment code, purpose, references, reclamation.
    Period decimal (109.66). Date DD.MM.YYYY.
    """

    bank_code = "530"
    bank_name = "NLB Banka"

    def __init__(self):
        super().__init__()
        self._bold_cid_map: dict[int, str] = {}
        self._regular_cid_map: dict[int, str] = {}

    def parse(self, file_path: Path) -> ParsedStatement:
        stmt = ParsedStatement(
            bank_code=self.bank_code,
            bank_name=self.bank_name,
        )

        # Build CID maps from embedded font outlines vs reference Arial
        self._bold_cid_map, self._regular_cid_map = _build_cid_maps_from_font(
            file_path
        )

        with pdfplumber.open(file_path) as pdf:
            if not pdf.pages:
                return stmt

            # Parse header from page 1 (bold text lines before tables)
            decoded_lines = self._decode_page(pdf.pages[0])
            self._parse_header(decoded_lines, pdf.pages[0], stmt)

            # Parse summary table (Table 0) using bold CID map
            self._parse_summary_from_table(pdf.pages[0], stmt)

            # Collect all transaction rows from all pages using table structure
            all_txn_rows = []
            for page in pdf.pages:
                page_rows = self._extract_table_rows(page)
                all_txn_rows.extend(page_rows)

            # Parse transactions
            self._parse_transaction_rows(all_txn_rows, stmt)

        return stmt

    # ------------------------------------------------------------------
    # Character decoding
    # ------------------------------------------------------------------

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
            return self._bold_cid_map.get(cid, "?"), True
        else:
            return self._regular_cid_map.get(cid, "?"), False

    def _decode_text(self, text: str) -> str:
        """Decode a CID-encoded string using the regular CID map.

        Input may contain (cid:XX) references mixed with regular text and
        newlines.  Each (cid:XX) is replaced with the decoded character.
        Unknown CIDs are replaced with '?'.
        """
        if not text:
            return ""

        def _replace_cid(m):
            cid = int(m.group(1))
            ch = self._regular_cid_map.get(cid)
            if ch is None:
                ch = self._bold_cid_map.get(cid, "?")
            return ch

        return re.sub(r"\(cid:(\d+)\)", _replace_cid, text)

    # ------------------------------------------------------------------
    # Cell / page decoding
    # ------------------------------------------------------------------

    def _decode_chars_in_bbox(self, page, bbox: tuple) -> str:
        """Decode all characters within a bounding box, using font-aware CID maps.

        Groups characters into lines (by y position) separated by newlines,
        with words separated by spaces (based on x-gap).
        Returns the decoded multi-line text.
        """
        if not bbox:
            return ""
        x0, y0, x1, y1 = bbox

        # Collect chars within bbox
        chars_in_cell = []
        for c in page.chars:
            cx, cy = c["x0"], c["top"]
            if x0 - 1 <= cx < x1 + 1 and y0 - 1 <= cy < y1 + 1:
                chars_in_cell.append(c)

        if not chars_in_cell:
            return ""

        # Group by y-line
        y_groups: dict[int, list] = defaultdict(list)
        for c in chars_in_cell:
            y_key = round(c["top"] / 4) * 4
            ch, is_bold = self._decode_char(c)
            if ch and isinstance(ch, str):
                y_groups[y_key].append((c["x0"], ch, c.get("width", 5)))

        # Build text with newlines between y-groups, spaces between words
        lines = []
        for y_key in sorted(y_groups.keys()):
            items = sorted(y_groups[y_key], key=lambda p: p[0])
            words = []
            current_word = ""
            current_end_x = None
            for x, ch, w in items:
                if current_end_x is not None and (x - current_end_x) > _SPACE_THRESHOLD:
                    if current_word:
                        words.append(current_word)
                    current_word = ch
                else:
                    current_word += ch
                current_end_x = x + max(w, 3)
            if current_word:
                words.append(current_word)
            lines.append(" ".join(words))

        return "\n".join(lines)

    def _decode_page(self, page) -> list[tuple[float, list[tuple[float, str]]]]:
        """Decode all characters on a page and group into lines of words.

        Returns list of (y_position, [(x_position, text), ...]) tuples.
        """
        chars = page.chars
        if not chars:
            return []

        decoded = []
        for c in chars:
            ch, is_bold = self._decode_char(c)
            if ch and isinstance(ch, str):
                decoded.append((c["x0"], c["top"], str(ch), is_bold, c.get("width", 5)))

        y_groups: dict[int, list] = defaultdict(list)
        for x, y, ch, is_bold, w in decoded:
            y_key = round(y / 4) * 4
            y_groups[y_key].append((x, ch, is_bold, w))

        result = []
        for y_key in sorted(y_groups.keys()):
            chars_on_line = sorted(y_groups[y_key], key=lambda p: p[0])
            words = []
            current_word = ""
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

    # ------------------------------------------------------------------
    # Header parsing
    # ------------------------------------------------------------------

    def _parse_header(self, lines: list, page, stmt: ParsedStatement) -> None:
        for y, words in lines:
            text = self._line_text(words)

            m = re.search(r"IZVOD\s*BR\.\s*(\d+)", text)
            if m:
                stmt.statement_number = m.group(1)

            m = re.search(r"DANA\s+(\d{2}\.\d{2}\.\d{4})", text)
            if m:
                stmt.statement_date = self.parse_date_dmy(m.group(1))

            if y < 120:
                for x, w in words:
                    if x < 200 and re.match(r"^[A-Z][A-Z\s]+$", w) and len(w) > 2:
                        if w not in ("IZVOD", "ZA", "STANJE", "NLB"):
                            stmt.client_name = self.clean_text(w)

            if y < 120:
                for x, w in words:
                    if re.match(r"530-\d{13}-\d{2}", w):
                        stmt.account_number = w

            m = re.search(r"poreski\s*broj\s*(\d+)", text, re.IGNORECASE)
            if m:
                stmt.client_pib = m.group(1)

    # ------------------------------------------------------------------
    # Summary parsing
    # ------------------------------------------------------------------

    def _parse_summary_from_table(self, page, stmt: ParsedStatement) -> None:
        """Parse Table 0 (summary) using find_tables() for cell bboxes
        and font-aware character decoding.

        Table 0 has 3 rows: header, sub-header, data.
        Data row cells: [opening_balance, total_debit, total_credit,
                        closing_balance, count_debit, count_credit]
        """
        tables = page.find_tables()
        if not tables:
            return

        t0 = tables[0]
        if len(t0.rows) < 3:
            return

        data_row = t0.rows[2]
        cells = data_row.cells

        decoded_cells = []
        for cell_bbox in cells:
            if cell_bbox is None:
                decoded_cells.append("")
            else:
                decoded_cells.append(self._decode_chars_in_bbox(page, cell_bbox))

        # Extract amounts from decoded cells
        amounts = []
        for cell in decoded_cells:
            m = re.search(r"([\d,]+\.\d{2})", cell)
            if m:
                amounts.append(m.group(1))
            else:
                amounts.append(None)

        # Expected: opening, total_debit, total_credit, closing, count_debit, count_credit
        if len(amounts) >= 4:
            if amounts[0]:
                stmt.opening_balance = self.parse_amount_us(amounts[0])
            if amounts[1]:
                stmt.total_debit = self.parse_amount_us(amounts[1])
            if amounts[2]:
                stmt.total_credit = self.parse_amount_us(amounts[2])
            if amounts[3]:
                stmt.closing_balance = self.parse_amount_us(amounts[3])

    # ------------------------------------------------------------------
    # Transaction table extraction
    # ------------------------------------------------------------------

    def _extract_table_rows(self, page) -> list[list[str]]:
        """Extract decoded transaction rows from a page.

        Finds the 9-column table (PROMJENE transactions), decodes each cell
        using font-aware character mapping, and returns decoded rows.
        Also handles "Ukupno" summary tables.
        """
        tables = page.find_tables()
        result_rows = []

        for table in tables:
            if not table.rows:
                continue

            # Count columns from the first row's cells
            first_row_cells = table.rows[0].cells
            ncols = len(first_row_cells)

            if ncols == 9:
                # This is the PROMJENE transactions table
                for row in table.rows:
                    decoded_row = []
                    for cell_bbox in row.cells:
                        if cell_bbox is None:
                            decoded_row.append("")
                        else:
                            decoded_row.append(
                                self._decode_chars_in_bbox(page, cell_bbox)
                            )
                    result_rows.append(decoded_row)

            elif ncols == 3:
                # Possibly "Ukupno EURA" summary row on page 2+
                for row in table.rows:
                    decoded_row = []
                    for cell_bbox in row.cells:
                        if cell_bbox is None:
                            decoded_row.append("")
                        else:
                            decoded_row.append(
                                self._decode_chars_in_bbox(page, cell_bbox)
                            )
                    # Mark as ukupno row by wrapping in special format
                    result_rows.append(["__UKUPNO__"] + decoded_row)

        return result_rows

    # ------------------------------------------------------------------
    # Transaction parsing
    # ------------------------------------------------------------------

    def _parse_transaction_rows(self, rows: list, stmt: ParsedStatement) -> None:
        """Parse transaction rows from the 9-column PROMJENE table.

        Rows 0-1 are headers, rows 2+ are data.
        Skip rows that look like headers or summary ("Ukupno") rows.
        """
        if not rows:
            return

        for row in rows:
            # Handle ukupno rows
            if row and row[0] == "__UKUPNO__":
                continue

            if not row or len(row) != 9:
                continue

            cells = row

            # Col 0: row number (single digit)
            col0 = cells[0].strip()
            if not col0 or not re.match(r"^\d+$", col0):
                continue

            row_num = int(col0)
            if row_num == 0:
                continue

            # Check for header-like rows
            full_text = " ".join(cells)
            if any(h in full_text.lower() for h in (
                "nal.", "naziv", "sjedi\u0161t", "\u0161ifra", "datum",
                "knji\u017een", "zadu\u017eenje", "odobrenje", "iznos",
                "svrha", "poziv", "reklamacij"
            )):
                continue

            txn = ParsedTransaction(row_number=row_num)

            # Col 1: counterparty + account
            self._parse_counterparty(cells[1], txn)

            # Col 2: origin + date
            self._parse_date(cells[2], txn)

            # Col 3: debit + fee
            self._parse_debit(cells[3], txn)

            # Col 4: credit
            self._parse_credit(cells[4], txn)

            # Col 5: payment code
            code = cells[5].strip()
            if code and re.match(r"^\d{3}$", code):
                txn.payment_code = code

            # Col 6: purpose
            purpose = cells[6].strip()
            if purpose:
                purpose = " ".join(purpose.split())
                txn.purpose = self.clean_text(purpose)

            # Col 7: references
            self._parse_references(cells[7], txn)

            # Col 8: reclamation
            reclam = cells[8].strip()
            if reclam:
                txn.reclamation_data = self.clean_text(reclam)

            if txn.debit is not None or txn.credit is not None:
                stmt.transactions.append(txn)

    def _parse_counterparty(self, text: str, txn: ParsedTransaction) -> None:
        """Parse counterparty name and account from col 1.

        Format: "COUNTERPARTY NAME, CITY,\\naccount-number"
        First line is name (take part before first comma), second line is account.
        """
        if not text:
            return

        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
        if not lines:
            return

        # First line(s): counterparty name
        # Last line might be account number (format: NNN-NNNNN...-NN)
        name_parts = []
        account = None
        for line in lines:
            acct_match = re.search(r"(\d{3}-[\d]+-\d{2})", line)
            if acct_match:
                account = acct_match.group(1)
            else:
                name_parts.append(line)

        if name_parts:
            full_name = " ".join(name_parts)
            parts = full_name.split(",")
            txn.counterparty = self.clean_text(parts[0])

        if account:
            txn.counterparty_account = self._normalize_account(account)

    def _normalize_account(self, account: str) -> str:
        """Normalize account: NNN-DIGITS-NN -> NNN-0000000DIGITS-NN (13 middle)."""
        m = re.match(r"^(\d{3})-(\d+)-(\d{2})$", account)
        if m:
            bank, mid, check = m.group(1), m.group(2), m.group(3)
            mid_padded = mid.zfill(13)
            return f"{bank}-{mid_padded}-{check}"
        return account

    def _parse_date(self, text: str, txn: ParsedTransaction) -> None:
        """Extract date DD.MM.YYYY from col 2 text."""
        if not text:
            return
        m = re.search(r"(\d{2}\.\d{2}\.\d{4})", text)
        if m:
            txn.booking_date = self.parse_date_dmy(m.group(1))
            txn.value_date = txn.booking_date

    def _parse_debit(self, text: str, txn: ParsedTransaction) -> None:
        """Parse debit amount and fee from col 3.

        Format: "amount\\nNaknada X.XX" or "-:-\\nNaknada 0.00"
        """
        if not text:
            return

        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
        if not lines:
            return

        # First line: debit amount
        first = lines[0]
        if first and first != "-:-" and "?" not in first:
            amt_match = re.search(r"([\d,]+\.\d{2})", first)
            if amt_match:
                amt = self.parse_amount_us(amt_match.group(1))
                if amt and amt != Decimal("0"):
                    txn.debit = amt

        # Look for Naknada (fee) in subsequent lines
        for line in lines[1:]:
            fee_match = re.search(r"[Nn]aknada\s*([\d,]+\.\d{2})", line)
            if fee_match:
                fee = self.parse_amount_us(fee_match.group(1))
                if fee and fee != Decimal("0"):
                    txn.fee = fee

    def _parse_credit(self, text: str, txn: ParsedTransaction) -> None:
        """Parse credit amount from col 4."""
        if not text:
            return
        text = text.strip()
        if not text or text == "-:-":
            return
        amt_match = re.search(r"([\d,]+\.\d{2})", text)
        if amt_match:
            amt = self.parse_amount_us(amt_match.group(1))
            if amt is not None and amt != Decimal("0"):
                txn.credit = amt

    def _parse_references(self, text: str, txn: ParsedTransaction) -> None:
        """Parse references from col 7.

        Format: "debit_ref\\ncredit_ref"
        """
        if not text:
            return
        lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
        if len(lines) >= 1 and lines[0]:
            txn.reference_debit = self.clean_text(lines[0])
        if len(lines) >= 2 and lines[1]:
            txn.reference_credit = self.clean_text(lines[1])
