from pathlib import Path
from typing import Optional

from app.parsers.base import BankParser, ParsedStatement

BANK_PARSERS: dict[str, type[BankParser]] = {}


def register_parser(parser_class: type[BankParser]) -> type[BankParser]:
    BANK_PARSERS[parser_class.bank_code] = parser_class
    return parser_class


def get_parser(bank_code: str) -> Optional[BankParser]:
    parser_class = BANK_PARSERS.get(bank_code)
    if parser_class:
        return parser_class()
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
