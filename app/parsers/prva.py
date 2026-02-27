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

            self._parse_header(full_text, stmt)
            self._parse_transactions(full_text, stmt)

        return stmt

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
