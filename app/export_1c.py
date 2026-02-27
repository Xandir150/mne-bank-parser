from datetime import datetime
from decimal import Decimal
from pathlib import Path

from app.models import Statement


def _fmt_date(d) -> str:
    """Format date as DD.MM.YYYY."""
    if d is None:
        return ""
    return d.strftime("%d.%m.%Y")


def _fmt_amount(val) -> str:
    """Format decimal amount with 2 decimal places using dot separator."""
    if val is None:
        return "0.00"
    return f"{Decimal(val):.2f}"


def generate_1c_file(statement: Statement, output_dir: Path) -> Path:
    """Generate a 1CClientBankExchange format file for the given statement.

    Args:
        statement: Statement ORM object with loaded transactions.
        output_dir: Directory to write the output file.

    Returns:
        Path to the generated file.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    date_str = statement.statement_date.strftime("%Y%m%d") if statement.statement_date else "nodate"
    filename = f"statement_{statement.id}_{statement.bank_code}_{date_str}.txt"
    output_path = output_dir / filename

    now = datetime.now()
    lines = []

    lines.append("1CClientBankExchange")
    lines.append("ВерсияФормата=1.03")
    lines.append("Кодировка=UTF-8")
    lines.append("Отправитель=BankStatementParser")
    lines.append("Получатель=")
    lines.append(f"ДатаСоздания={now.strftime('%d.%m.%Y')}")
    lines.append(f"ВремяСоздания={now.strftime('%H:%M:%S')}")
    lines.append(f"ДатаНачала={_fmt_date(statement.period_start)}")
    lines.append(f"ДатаКонца={_fmt_date(statement.period_end)}")
    lines.append(f"РасчСчет={statement.account_number or ''}")

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
            # Debit: payer = our company, receiver = counterparty
            lines.append(f"ПлательщикСчет={statement.account_number or ''}")
            lines.append(f"Плательщик={statement.client_name or ''}")
            lines.append(f"ПлательщикИНН={statement.client_pib or ''}")
            lines.append(f"ПолучательСчет={tx.counterparty_account or ''}")
            lines.append(f"Получатель={tx.counterparty or ''}")
            lines.append("ПолучательИНН=")
        else:
            # Credit: payer = counterparty, receiver = our company
            lines.append(f"ПлательщикСчет={tx.counterparty_account or ''}")
            lines.append(f"Плательщик={tx.counterparty or ''}")
            lines.append("ПлательщикИНН=")
            lines.append(f"ПолучательСчет={statement.account_number or ''}")
            lines.append(f"Получатель={statement.client_name or ''}")
            lines.append(f"ПолучательИНН={statement.client_pib or ''}")

        lines.append(f"НазначениеПлатежа={tx.purpose or ''}")
        lines.append("КонецДокумента")

    lines.append("КонецФайла")

    content = "\r\n".join(lines)
    output_path.write_text(content, encoding="utf-8")

    return output_path
