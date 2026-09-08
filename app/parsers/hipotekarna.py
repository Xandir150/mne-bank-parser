import logging
import re
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional

import pdfplumber

from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement, ParsedTransaction

logger = logging.getLogger(__name__)

# Above this, an amount is physically implausible for a bank statement line
# and is almost certainly a column-misclassification artifact (e.g. a
# reference-number digit string parsed as an amount). See _parse_page.
_MAX_PLAUSIBLE_AMOUNT = Decimal("10000000")


@register_parser
class HipotekarnaParser(BankParser):
    """Parser for Hipotekarna Banka (520) PDF statements.

    Two known export layouts:
    - Single-statement export: text-only PDF (no vector lines). Header has
      client/bank info, then IZVOD BR. line, then transaction rows, then
      summary. Parsed via word x-position (_parse_page_words).
    - Bulk "print whole period" export (one page per statement day): has
      real vector table lines, so pages are parsed structurally via
      pdfplumber's lattice extract_tables() (_parse_page_lattice). Note
      parse() folds every embedded day's transactions into a single
      ParsedStatement, matching the single-statement layout's shape —
      BankParser.parse_multi() would be the more semantically correct API
      for this bundle format (each page is its own dated mini-statement),
      but isn't implemented here; see hipotekarna.py notes.

    Both layouts share the same 8-column table: Valuta, Counterparty+account,
    Bank, Debit, Credit, Payment code+purpose, References, Reclamation —
    every page (in either layout) prints a "1".."8" marker row directly
    above the transaction table, which both parsing strategies use to
    locate/validate the table.
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
                # PIB is a short numeric at the right edge of the client-name
                # line. Take the first match only: reference numbers in the
                # transaction table below sit at a similar x and can also be
                # 7-8 digits (e.g. "13082026"), and would otherwise overwrite it.
                if re.match(r"^\d{7,8}$", t) and x > 1000 and not stmt.client_pib:
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

    # Fallback column x-boundaries, used only when a page doesn't carry the
    # "1".."8" column-marker row that _find_column_boundaries looks for
    # (see below). Calibrated against the 1755x1240pt single-statement
    # export layout ("good_single.pdf" in bug reports) — that layout has no
    # marker row, so this is what it actually runs on.
    # x < 150: date (col 1)
    # 150 <= x < 600: counterparty name / account (col 2)
    # 600 <= x < 800: debit amount (col 4)
    # 800 <= x < 880: credit amount (col 5)
    # 880 <= x < 1200: payment code / purpose (col 6)
    # 1200 <= x < 1450: references (col 7)
    # x >= 1450: reclamation data (col 8)
    _DEFAULT_BOUNDS = (150, 600, 750, 880, 1200, 1450)

    # A line consisting of exactly these 8 tokens, in order, is the
    # numbered column-header row Hipotekarna prints above the transaction
    # table on most (but not all — pages with zero transactions omit the
    # table entirely) pages.
    _MARKER_TOKENS = [str(n) for n in range(1, 9)]

    def _find_column_boundaries(self, lines) -> Optional[tuple]:
        """Locate the '1'..'8' column-marker row on a page and derive

        column boundaries from it (midpoints between consecutive markers).
        This makes column classification scale/DPI-invariant instead of
        relying on pixel constants tuned for one specific page size.
        Returns a 6-tuple (col_counterparty, col_debit, col_credit,
        col_purpose, col_ref, col_reclam) boundary, or None if no such row
        is found on this page (caller should fall back to _DEFAULT_BOUNDS).
        """
        for y, ws in lines:
            texts = [w['text'].strip() for w in ws if w['text'].strip()]
            if texts == self._MARKER_TOKENS:
                xs = [w['x0'] for w in ws if w['text'].strip()]
                return (
                    (xs[0] + xs[1]) / 2,  # before col 2: counterparty+account
                    (xs[2] + xs[3]) / 2,  # before col 4: debit
                    (xs[3] + xs[4]) / 2,  # before col 5: credit
                    (xs[4] + xs[5]) / 2,  # before col 6: payment code/purpose
                    (xs[5] + xs[6]) / 2,  # before col 7: references
                    (xs[6] + xs[7]) / 2,  # before col 8: reclamation
                )
        return None

    def _classify_word(self, x: float, bounds: tuple) -> str:
        col_counterparty, col_debit, col_credit, col_purpose, col_ref, col_reclam = bounds
        if x < col_counterparty:
            return "date"
        elif x < col_debit:
            return "counterparty"
        elif x < col_credit:
            return "debit"
        elif x < col_purpose:
            return "credit"
        elif x < col_ref:
            return "purpose"
        elif x < col_reclam:
            return "reference"
        else:
            return "reclamation"

    def _validate_transaction(self, txn: ParsedTransaction, date_str: str) -> None:
        """Defense-in-depth sanity check.

        A physically implausible amount, or a counterparty field that looks
        like it swallowed the debit/credit/payment-code columns, is a strong
        signal that column classification went wrong somewhere (e.g. a
        reference-number digit string misread as an amount) — regardless of
        whether it's the dynamic marker-row detection or the hardcoded
        fallback that produced the boundaries. Rather than silently export
        bad numbers into 1C (as happened in production before this check
        existed), raise loudly so the whole file is held back — see
        worker._process_file, which catches this, marks the statement
        status="error", and leaves the file in place for a human to review
        and retry.
        """
        for label, amount in (("debit", txn.debit), ("credit", txn.credit)):
            if amount is not None and amount > _MAX_PLAUSIBLE_AMOUNT:
                raise ValueError(
                    f"Hipotekarna parser: implausible {label} amount {amount} on "
                    f"{date_str} (counterparty={txn.counterparty!r}) — likely column "
                    f"misclassification (e.g. a reference number parsed as an "
                    f"amount); aborting parse rather than exporting bad data"
                )

        cp = txn.counterparty or ""
        decimal_numbers = re.findall(r"\d+[.,]\d{2}\b", cp)
        payment_code_like = re.search(r"\b\d{3}\b", cp)
        if len(decimal_numbers) >= 2 and payment_code_like:
            raise ValueError(
                f"Hipotekarna parser: counterparty field on {date_str} looks like "
                f"it swallowed amount/payment-code columns ({cp!r}) — likely "
                f"column misclassification; aborting parse rather than exporting "
                f"bad data"
            )

    def _parse_page(self, page, stmt: ParsedStatement) -> None:
        """Dispatch to the appropriate extraction strategy for this page.

        Hipotekarna's bulk "print whole period" export (one page per
        statement day) renders the transaction table with real vector
        table lines — pdfplumber's lattice `extract_tables()` parses those
        pages perfectly and structurally (no column-position guessing at
        all). The older single-statement export has no vector lines at
        all (text-only layout), so it can't use lattice extraction and
        needs the word-position-based fallback below.
        """
        if page.lines:
            self._parse_page_lattice(page, stmt)
        else:
            self._parse_page_words(page, stmt)

    def _parse_page_lattice(self, page, stmt: ParsedStatement) -> None:
        """Parse a page's transaction table via pdfplumber's lattice extraction.

        Structurally identifies the transaction table by locating the
        "1".."8" marker row within it (same anchor used by the word-based
        fallback) and treats every row after it as a transaction, using
        fixed column indices (0=date, 1=name+account, 2=bank,
        3=debit, 4=credit, 5=payment code+purpose, 6=reference,
        7=reclamation) — those indices are a structural property of the
        table itself, not pixel positions, so they don't need per-page
        calibration. Also parses the adjacent daily-summary table for
        opening/closing balance and turnover totals.
        """
        tables = page.extract_tables()
        txn_rows: list = []
        summary_rows: Optional[list] = None

        for table in tables:
            marker_idx = None
            for i, row in enumerate(table):
                norm = [(c or "").strip() for c in row]
                if norm == self._MARKER_TOKENS:
                    marker_idx = i
                    break
            if marker_idx is not None:
                txn_rows = table[marker_idx + 1:]
                continue
            if table and table[0] and (table[0][0] or "").strip() == "Previous state":
                summary_rows = table

        if not txn_rows:
            logger.warning(
                "Hipotekarna page %s: has vector table lines but no '1'..'8' "
                "marker row was found in any extracted table — skipping "
                "transaction extraction for this page",
                getattr(page, "page_number", "?"),
            )

        for row in txn_rows:
            date_text = (row[0] or "").strip()
            if not re.match(r"\d{1,2}\.\d{1,2}\.\d{4}\.?$", date_text):
                continue  # not a transaction row (e.g. a trailing blank row)

            name_lines = (row[1] or "").split("\n")
            counterparty = self.clean_text(name_lines[0]) if name_lines else ""
            counterparty_account = (
                self.clean_text(name_lines[1]) if len(name_lines) > 1 else ""
            )

            debit = self.parse_amount_us((row[3] or "").strip())
            credit = self.parse_amount_us((row[4] or "").strip())
            if debit == Decimal("0"):
                debit = None
            if credit == Decimal("0"):
                credit = None
            if debit is None and credit is None:
                continue

            code_lines = (row[5] or "").split("\n")
            code_line = code_lines[0].strip() if code_lines else ""
            code_match = re.match(r"^(\d{3})\b", code_line)
            payment_code = code_match.group(1) if code_match else ""
            if len(code_lines) > 1:
                purpose = self.clean_text(code_lines[1])
            else:
                purpose = self.clean_text(re.sub(r"^\d+\s*", "", code_line))

            ref_lines = (row[6] or "").split("\n")
            reference_debit = self.clean_text(ref_lines[0]) if ref_lines and ref_lines[0].strip() else None
            reference_credit = self.clean_text(ref_lines[1]) if len(ref_lines) > 1 and ref_lines[1].strip() else None

            reclamation_data = self.clean_text(row[7]) if row[7] else None

            txn = ParsedTransaction(row_number=len(stmt.transactions) + 1)
            txn.value_date = self.parse_date_dmy(date_text)
            txn.booking_date = txn.value_date
            txn.counterparty = counterparty
            txn.counterparty_account = counterparty_account
            txn.payment_code = payment_code
            txn.purpose = purpose
            txn.debit = debit
            txn.credit = credit
            if reference_debit:
                txn.reference_debit = reference_debit
            if reference_credit:
                txn.reference_credit = reference_credit
            if reclamation_data:
                txn.reclamation_data = reclamation_data

            self._validate_transaction(txn, date_text)
            stmt.transactions.append(txn)

        if summary_rows and len(summary_rows) >= 2:
            vals = summary_rows[1]
            opening = self.parse_amount_us(vals[0]) if len(vals) > 0 else None
            debit_total = self.parse_amount_us(vals[1]) if len(vals) > 1 else None
            credit_total = self.parse_amount_us(vals[2]) if len(vals) > 2 else None
            closing = self.parse_amount_us(vals[3]) if len(vals) > 3 else None

            if stmt.opening_balance is None:
                stmt.opening_balance = opening
            if closing is not None:
                stmt.closing_balance = closing  # last page processed wins
            if debit_total is not None:
                stmt.total_debit = (stmt.total_debit or Decimal("0")) + debit_total
            if credit_total is not None:
                stmt.total_credit = (stmt.total_credit or Decimal("0")) + credit_total

    def _parse_page_words(self, page, stmt: ParsedStatement) -> None:
        """Fallback parser for pages with no vector table lines (lattice

        extraction isn't possible there) — classifies words into columns
        by x-position, using dynamic per-page boundaries anchored on the
        "1".."8" marker row when present, falling back to hardcoded
        pixel constants (_DEFAULT_BOUNDS) otherwise.
        """
        words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)
        lines = self._group_by_y(words, tolerance=4)

        bounds = self._find_column_boundaries(lines)
        if bounds is None:
            bounds = self._DEFAULT_BOUNDS
            # Pages with no transactions at all (e.g. a "no activity that
            # day" page in a bulk multi-day export) legitimately have no
            # table and thus no marker row — the fallback is harmless there
            # since there's no data to misclassify. Only warn when the page
            # actually looks like it has transaction rows to parse.
            has_txn_lines = any(
                re.match(r"\d{2}\.\d{2}\.\d{4}", " ".join(w['text'].strip() for w in ws))
                for _, ws in lines
            )
            if has_txn_lines:
                logger.warning(
                    "Hipotekarna page %s: no '1'..'8' column-marker row found — "
                    "falling back to hardcoded column boundaries %s",
                    getattr(page, "page_number", "?"), bounds,
                )

        # Pair lines: line1 (date row) + line2 (account row)
        pending_txn = None

        for y, ws in lines:
            cols = {}
            for w in ws:
                text = w['text'].strip()
                if not text:
                    continue
                col = self._classify_word(w['x0'], bounds)
                cols.setdefault(col, []).append(text)

            date_text = " ".join(cols.get("date", []))

            # Check if this is a summary line (all numbers, no date pattern)
            if not re.match(r"\d{2}\.\d{2}\.\d{4}", date_text):
                # Could be line 2 of a transaction or summary
                cp_text = " ".join(cols.get("counterparty", []))

                if pending_txn and re.match(r"(\d{6,18}|\d{3}-\d{5}-\d{2})", cp_text):
                    # Line 2: account or branch code, purpose, reference, reclamation
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

            self._validate_transaction(txn, date_str)

            pending_txn = txn

        # Flush last pending
        if pending_txn and (pending_txn.debit is not None or pending_txn.credit is not None):
            stmt.transactions.append(pending_txn)
