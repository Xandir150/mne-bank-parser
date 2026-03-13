import logging
import re
from pathlib import Path
from typing import Optional

from app.parsers.base import BankParser, ParsedStatement

logger = logging.getLogger(__name__)

BANK_PARSERS: dict[str, type[BankParser]] = {}

# Text markers for bank detection (case-insensitive, checked in order)
_BANK_TEXT_MARKERS: list[tuple[str, list[str]]] = [
    ("520", ["hipotekarna"]),
    ("530", ["nlb banka", "nlb "]),
    ("535", ["prva banka"]),
    ("540", ["erste", "windows-1250"]),
    ("560", ["universal capital", "broj partije"]),
    ("565", ["lovćen", "lovcen"]),
    ("570", ["zapad", "выписка по счету"]),
    ("575", ["ziraat"]),
    ("580", ["adriatic", "statement turnover"]),
]


def register_parser(parser_class: type[BankParser]) -> type[BankParser]:
    BANK_PARSERS[parser_class.bank_code] = parser_class
    return parser_class


def get_parser(bank_code: str) -> Optional[BankParser]:
    parser_class = BANK_PARSERS.get(bank_code)
    if parser_class:
        return parser_class()
    return None


def detect_bank_code(file_path: Path) -> Optional[str]:
    """Auto-detect bank code from file content.

    Detection priority:
    1. IBAN pattern ME25NNN... → bank code NNN
    2. Account number pattern NNN-... → bank code NNN
    3. Text markers (bank name mentions)
    """
    text = _extract_first_page_text(file_path)
    if not text:
        # Fallback: try to detect bank code from filename
        return _detect_from_filename(file_path)

    known_codes = set(BANK_PARSERS.keys())

    # Normalize: strip spaces from IBANs for matching
    text_no_iban_spaces = re.sub(r"(ME\d{2})\s+", r"\1", text)

    # 1. IBAN: ME25NNN...
    m = re.search(r"ME\d{2}(\d{3})", text_no_iban_spaces)
    if m and m.group(1) in known_codes:
        return m.group(1)

    # 2. Account number: NNN-digits-NN or NNN followed by 15+ digits
    m = re.search(r"\b(\d{3})-\d+-\d{2}\b", text)
    if m and m.group(1) in known_codes:
        return m.group(1)
    m = re.search(r"\b(\d{3})\d{15,}\b", text)
    if m and m.group(1) in known_codes:
        return m.group(1)

    # 3. Text markers (case-insensitive)
    text_lower = text.lower()
    for code, markers in _BANK_TEXT_MARKERS:
        for marker in markers:
            if marker in text_lower:
                return code

    # 4. Fallback: filename
    return _detect_from_filename(file_path)


def _detect_from_filename(file_path: Path) -> Optional[str]:
    """Try to detect bank code from filename patterns like '530-' or '520000000'."""
    name = file_path.name
    known_codes = set(BANK_PARSERS.keys())
    # Pattern: NNN- at start or after common separators
    m = re.search(r"(\d{3})-\d", name)
    if m and m.group(1) in known_codes:
        return m.group(1)
    # Pattern: NNN followed by many digits (account number in filename)
    m = re.search(r"(\d{3})\d{12,}", name)
    if m and m.group(1) in known_codes:
        return m.group(1)
    return None


def _extract_first_page_text(file_path: Path) -> Optional[str]:
    """Extract text from the first page of a PDF or full HTML content."""
    suffix = file_path.suffix.lower()
    if suffix in (".htm", ".html"):
        try:
            # Try common encodings for Erste Bank HTML files
            for enc in ("utf-8", "windows-1250", "windows-1251"):
                try:
                    return file_path.read_text(encoding=enc)[:3000]
                except UnicodeDecodeError:
                    continue
        except Exception:
            return None
    elif suffix == ".pdf":
        try:
            import pdfplumber
            with pdfplumber.open(file_path) as pdf:
                if pdf.pages:
                    return (pdf.pages[0].extract_text() or "")[:3000]
        except Exception:
            return None
    return None


def parse_file(file_path: Path, bank_code: str) -> ParsedStatement:
    parser = get_parser(bank_code)
    if not parser:
        raise ValueError(f"No parser registered for bank code: {bank_code}")
    return parser.parse(file_path)


def get_registered_banks() -> dict[str, str]:
    return {code: cls.bank_name for code, cls in BANK_PARSERS.items()}


# Import all parsers to trigger registration
def _load_parsers():
    import importlib
    import pkgutil
    package_path = Path(__file__).parent
    for _, module_name, _ in pkgutil.iter_modules([str(package_path)]):
        if module_name != "base":
            importlib.import_module(f"app.parsers.{module_name}")


_load_parsers()
