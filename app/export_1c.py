import re
from datetime import datetime
from decimal import Decimal
from pathlib import Path

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

    Output: output/{account}/{date}_{statement_id}.txt
    Each statement gets its own file. 1C "Файл загрузки из банка" points
    to the account directory — user selects the specific file to import.
    """
    account_dir_name = _fmt_account(statement.account_number) or "unknown"
    account_dir = output_dir / account_dir_name
    account_dir.mkdir(parents=True, exist_ok=True)

    date_str = statement.statement_date.strftime("%Y%m%d") if statement.statement_date else "nodate"
    filename = f"{date_str}_{statement.id}.txt"
    output_path = account_dir / filename

    stmt_date = statement.period_start or statement.statement_date
    stmt_end = statement.period_end or statement.statement_date

    now = datetime.now()
    lines = []

    lines.append("1CClientBankExchange")
    lines.append("ВерсияФормата=1.03")
    lines.append("Кодировка=Windows")
    lines.append("Отправитель=BankStatementParser")
    lines.append("Получатель=")
    lines.append(f"ДатаСоздания={now.strftime('%d.%m.%Y')}")
    lines.append(f"ВремяСоздания={now.strftime('%H:%M:%S')}")
    lines.append(f"ДатаНачала={_fmt_date(stmt_date)}")
    lines.append(f"ДатаКонца={_fmt_date(stmt_end)}")
    lines.append(f"РасчСчет={_fmt_account(statement.account_number)}")

    # Account section
    lines.append("СекцияРасчСчет")
    lines.append(f"НачальныйОстаток={_fmt_amount(statement.opening_balance)}")
    lines.append(f"КонечныйОстаток={_fmt_amount(statement.closing_balance)}")
    lines.append(f"ДебетОборот={_fmt_amount(statement.total_debit)}")
    lines.append(f"КредитОборот={_fmt_amount(statement.total_credit)}")
    lines.append("КонецРасчСчет")

    # Transactions
    for tx in statement.transactions:
        is_debit = tx.debit is not None and tx.debit > 0
        amount = tx.debit if is_debit else tx.credit
        tx_date = tx.value_date or tx.booking_date

        lines.append("СекцияДокумент=Платёжное поручение")
        lines.append(f"Номер={tx.row_number}")
        lines.append(f"Дата={_fmt_date(tx_date)}")
        lines.append(f"Сумма={_fmt_amount(amount)}")

        if is_debit:
            lines.append(f"ПлательщикСчет={_fmt_account(statement.account_number)}")
            lines.append(f"Плательщик={_safe_text(statement.client_name or '')}")
            lines.append(f"ПлательщикИНН={statement.client_pib or ''}")
            lines.append(f"ПолучательСчет={_fmt_account(tx.counterparty_account)}")
            lines.append(f"Получатель={_safe_text(tx.counterparty or '')}")
            lines.append("ПолучательИНН=")
        else:
            lines.append(f"ПлательщикСчет={_fmt_account(tx.counterparty_account)}")
            lines.append(f"Плательщик={_safe_text(tx.counterparty or '')}")
            lines.append("ПлательщикИНН=")
            lines.append(f"ПолучательСчет={_fmt_account(statement.account_number)}")
            lines.append(f"Получатель={_safe_text(statement.client_name or '')}")
            lines.append(f"ПолучательИНН={statement.client_pib or ''}")

        lines.append(f"НазначениеПлатежа={_safe_text(tx.purpose or '')}")
        lines.append("КонецДокумента")

    lines.append("КонецФайла")

    content = "\r\n".join(lines)
    output_path.write_text(content, encoding="windows-1251")

    return output_path
