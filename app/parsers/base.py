from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import date
from decimal import Decimal
from pathlib import Path
from typing import Optional
import re


@dataclass
class ParsedTransaction:
    row_number: int = 0
    value_date: Optional[date] = None
    booking_date: Optional[date] = None
    debit: Optional[Decimal] = None
    credit: Optional[Decimal] = None
    counterparty: Optional[str] = None
    counterparty_account: Optional[str] = None
    counterparty_bank: Optional[str] = None
    payment_code: Optional[str] = None
    purpose: Optional[str] = None
    reference_debit: Optional[str] = None
    reference_credit: Optional[str] = None
    reclamation_data: Optional[str] = None
    fee: Optional[Decimal] = None


@dataclass
class ParsedStatement:
    bank_code: str = ""
    bank_name: str = ""
    account_number: str = ""
    iban: Optional[str] = None
    statement_number: Optional[str] = None
    statement_date: Optional[date] = None
    period_start: Optional[date] = None
    period_end: Optional[date] = None
    opening_balance: Optional[Decimal] = None
    closing_balance: Optional[Decimal] = None
    total_debit: Optional[Decimal] = None
    total_credit: Optional[Decimal] = None
    currency: str = "EUR"
    client_name: Optional[str] = None
    client_pib: Optional[str] = None
    transactions: list[ParsedTransaction] = field(default_factory=list)


class BankParser(ABC):
    bank_code: str = ""
    bank_name: str = ""

    @abstractmethod
    def parse(self, file_path: Path) -> ParsedStatement:
        """Parse a bank statement file and return structured data."""
        ...

    # --- Utility methods for subclasses ---

    @staticmethod
    def parse_amount_eu(text: str) -> Optional[Decimal]:
        """Parse European format: 1.234,56 -> 1234.56"""
        if not text or not text.strip():
            return None
        text = text.strip().replace(" ", "")
        if not text or text == "-":
            return None
        text = text.replace(".", "").replace(",", ".")
        try:
            return Decimal(text)
        except Exception:
            return None

    @staticmethod
    def parse_amount_us(text: str) -> Optional[Decimal]:
        """Parse US/international format: 1,234.56 -> 1234.56"""
        if not text or not text.strip():
            return None
        text = text.strip().replace(" ", "")
        if not text or text == "-":
            return None
        text = text.replace(",", "")
        try:
            return Decimal(text)
        except Exception:
            return None

    @staticmethod
    def parse_date_dmy(text: str) -> Optional[date]:
        """Parse DD.MM.YYYY or DD.MM.YYYY. format"""
        if not text or not text.strip():
            return None
        text = text.strip().rstrip(".")
        m = re.match(r"(\d{1,2})[./](\d{1,2})[./](\d{4})", text)
        if m:
            try:
                return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
        return None

    @staticmethod
    def parse_date_ymd(text: str) -> Optional[date]:
        """Parse YYYY.MM.DD format"""
        if not text or not text.strip():
            return None
        text = text.strip().rstrip(".")
        m = re.match(r"(\d{4})[./](\d{1,2})[./](\d{1,2})", text)
        if m:
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None
        return None

    @staticmethod
    def parse_date_dmy_slash(text: str) -> Optional[date]:
        """Parse DD/MM/YYYY format"""
        if not text or not text.strip():
            return None
        text = text.strip()
        m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", text)
        if m:
            try:
                return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
        return None

    @staticmethod
    def clean_text(text: str) -> str:
        """Clean whitespace from parsed text."""
        if not text:
            return ""
        return " ".join(text.split()).strip()
