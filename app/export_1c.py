import re
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import List

from app.models import Statement


# Serbian/Croatian Latin characters not present in Windows-1251
_LATIN_MAP = str.maketrans({
    'š': 's', 'Š': 'S',
    'č': 'c', 'Č': 'C',
    'ć': 'c', 'Ć': 'C',
    'ž': 'z', 'Ž': 'Z',
    'đ': 'dj', 'Đ': 'Dj',
})


def _safe_text(text: str) -> str:
    """Replace characters that cannot be encoded in Windows-1251."""
    if not text:
        return ""
    return text.translate(_LATIN_MAP)


def _fmt_date(d) -> str:
    """Format date as DD.MM.YYYY."""
    if d is None:
        return ""
    return d.strftime("%d.%m.%Y")


def _fmt_account(acct) -> str:
    """Format account number for 1C: 18-digit format without dashes.

    Montenegrin accounts have format: BBB-NNNNNNNNNNNNN-CC (3-13-2).
    Short formats like 535-22023-67 need zero-padding in the middle part.
    """
    if not acct:
        return ""
    m = re.match(r"^(\d{3})-(\d+)-(\d{2})$", acct)
    if m:
        bank, number, check = m.group(1), m.group(2), m.group(3)
        return bank + number.zfill(13) + check
    # Already without dashes or unknown format
    return acct.replace("-", "")


def _fmt_amount(val) -> str:
    """Format decimal amount with 2 decimal places using dot separator."""
    if val is None:
        return "0.00"
    return f"{Decimal(val):.2f}"


def generate_1c_file(statement: Statement, output_dir: Path) -> Path:
    """Generate a 1CClientBankExchange file for a single statement.

    The file is written to output/{account}/import.txt.
    Each new export overwrites the file for this account.
    1C deduplicates by date+number so previously imported operations won't repeat.
    """
    return generate_1c_file_multi([statement], output_dir)


def generate_1c_file_multi(statements: List[Statement], output_dir: Path) -> Path:
    """Generate a single 1CClientBankExchange file containing all given statements.

    All statements must belong to the same account.
    File is written to output/{account}/import.txt.
    """
    if not statements:
        raise ValueError("No statements to export")

    account_number = statements[0].account_number
    account_dir_name = _fmt_account(account_number) or "unknown"
    account_dir = output_dir / account_dir_name
    account_dir.mkdir(parents=True, exist_ok=True)
    output_path = account_dir / "import.txt"

    # Determine date range across all statements
    all_dates = []
    for stmt in statements:
        if stmt.period_start:
            all_dates.append(stmt.period_start)
        if stmt.period_end:
            all_dates.append(stmt.period_end)
        if stmt.statement_date:
            all_dates.append(stmt.statement_date)

    date_start = min(all_dates) if all_dates else None
    date_end = max(all_dates) if all_dates else None

    now = datetime.now()
    lines = []

    # File header
    lines.append("1CClientBankExchange")
    lines.append("ВерсияФормата=1.03")
    lines.append("Кодировка=Windows")
    lines.append("Отправитель=BankStatementParser")
    lines.append("Получатель=")
    lines.append(f"ДатаСоздания={now.strftime('%d.%m.%Y')}")
    lines.append(f"ВремяСоздания={now.strftime('%H:%M:%S')}")
    lines.append(f"ДатаНачала={_fmt_date(date_start)}")
    lines.append(f"ДатаКонца={_fmt_date(date_end)}")
    lines.append(f"РасчСчет={_fmt_account(account_number)}")

    # One account section per statement
    for stmt in statements:
        lines.append("СекцияРасчСчет")
        lines.append(f"ДатаНачала={_fmt_date(stmt.period_start or stmt.statement_date)}")
        lines.append(f"ДатаКонца={_fmt_date(stmt.period_end or stmt.statement_date)}")
        lines.append(f"НачальныйОстаток={_fmt_amount(stmt.opening_balance)}")
        lines.append(f"КонечныйОстаток={_fmt_amount(stmt.closing_balance)}")
        lines.append(f"ДебетОборот={_fmt_amount(stmt.total_debit)}")
        lines.append(f"КредитОборот={_fmt_amount(stmt.total_credit)}")
        lines.append("КонецРасчСчет")

    # All transactions from all statements
    for stmt in statements:
        for tx in stmt.transactions:
            is_debit = tx.debit is not None and tx.debit > 0
            amount = tx.debit if is_debit else tx.credit
            tx_date = tx.value_date or tx.booking_date

            lines.append("СекцияДокумент=Платёжное поручение")
            lines.append(f"Номер={tx.row_number}")
            lines.append(f"Дата={_fmt_date(tx_date)}")
            lines.append(f"Сумма={_fmt_amount(amount)}")

            if is_debit:
                lines.append(f"ПлательщикСчет={_fmt_account(stmt.account_number)}")
                lines.append(f"Плательщик={_safe_text(stmt.client_name or '')}")
                lines.append(f"ПлательщикИНН={stmt.client_pib or ''}")
                lines.append(f"ПолучательСчет={_fmt_account(tx.counterparty_account)}")
                lines.append(f"Получатель={_safe_text(tx.counterparty or '')}")
                lines.append("ПолучательИНН=")
            else:
                lines.append(f"ПлательщикСчет={_fmt_account(tx.counterparty_account)}")
                lines.append(f"Плательщик={_safe_text(tx.counterparty or '')}")
                lines.append("ПлательщикИНН=")
                lines.append(f"ПолучательСчет={_fmt_account(stmt.account_number)}")
                lines.append(f"Получатель={_safe_text(stmt.client_name or '')}")
                lines.append(f"ПолучательИНН={stmt.client_pib or ''}")

            lines.append(f"НазначениеПлатежа={_safe_text(tx.purpose or '')}")
            lines.append("КонецДокумента")

    lines.append("КонецФайла")

    content = "\r\n".join(lines)
    output_path.write_text(content, encoding="windows-1251")

    return output_path
