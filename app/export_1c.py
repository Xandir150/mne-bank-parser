import re
from datetime import datetime
from decimal import Decimal
from pathlib import Path

from app.models import Statement


# Montenegrin payment codes (šifra plaćanja) → Russian descriptions for 1C
# Structure: 1xx = cash, 2xx = non-cash; last two digits = payment basis
PAYMENT_CODE_DESCRIPTIONS = {
    # Товары и услуги
    "120": "Товары", "220": "Товары",
    "121": "Услуги", "221": "Комиссия банка",
    "122": "Товары", "222": "Товары",
    "123": "Инвестиции", "223": "Инвестиции",
    "124": "Инвестиции прочие", "224": "Инвестиции прочие",
    "125": "Аренда гос. имущества", "225": "Аренда гос. имущества",
    "126": "Аренда", "226": "Аренда",
    "127": "Субсидии", "227": "Субсидии",
    "128": "Субсидии прочие", "228": "Субсидии прочие",
    # Таможня
    "131": "Таможенные пошлины", "231": "Таможенные пошлины",
    "132": "Акцизы и сборы", "232": "Акцизы и сборы",
    # Налоги и зарплата
    "139": "Прирез на подоходный налог", "239": "Прирез на подоходный налог",
    "140": "Налоги и взносы", "240": "Заработная плата",
    "141": "Необлагаемые выплаты", "241": "Необлагаемые выплаты",
    "142": "Компенсации по зарплате", "242": "Компенсации по зарплате",
    "145": "Пенсии", "245": "Пенсии",
    "146": "Удержания", "246": "Удержания",
    "148": "Доходы от капитала", "248": "Доходы от капитала",
    "151": "Заработная плата", "251": "Заработная плата",
    "153": "Налоги и сборы", "253": "Налоги и сборы",
    "154": "Коммунальные услуги", "254": "Коммунальные услуги",
    "157": "Страховые взносы", "257": "Страховые взносы",
    "158": "Членские взносы", "258": "Членские взносы",
    # Переводы
    "160": "Премии", "260": "Премии",
    "161": "Ценные бумаги", "261": "Ценные бумаги",
    "162": "Трансфертные платежи", "262": "Трансфертные платежи",
    "163": "Прочие переводы", "263": "Прочие переводы",
    "165": "Выручка", "265": "Выручка",
    "166": "Снятие наличных", "266": "Снятие наличных",
    # Кредиты и финансы
    "170": "Краткосрочные кредиты", "270": "Краткосрочные кредиты",
    "171": "Долгосрочные кредиты", "271": "Долгосрочные кредиты",
    "172": "Проценты по кредитам", "272": "Проценты по кредитам",
    "173": "Проценты", "273": "Проценты",
    "177": "Внутренние расчеты банка", "277": "Внутренние расчеты банка",
    "178": "Купля-продажа валюты", "278": "Купля-продажа валюты",
    # Прочие
    "185": "Пожертвования", "285": "Пожертвования",
    "186": "Выплата ущерба", "286": "Выплата ущерба",
    "187": "Платежи за третьих лиц", "287": "Платежи за третьих лиц",
    "189": "Прочие платежи", "289": "Прочие платежи",
    "190": "Валютные операции", "290": "Валютные операции",
}


def _purpose_with_code(payment_code, purpose: str) -> str:
    """Prepend Russian payment code description to purpose text."""
    if not payment_code:
        return purpose or ""
    desc = PAYMENT_CODE_DESCRIPTIONS.get(payment_code)
    if desc:
        prefix = f"{payment_code} {desc}"
    else:
        prefix = f"Код {payment_code}"
    if purpose:
        return f"[{prefix}] {purpose}"
    return f"[{prefix}]"


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


def _safe_dirname(name: str) -> str:
    """Sanitize a string for use as a directory name."""
    if not name:
        return ""
    # Transliterate Serbian characters
    name = name.translate(_LATIN_MAP)
    # Replace filesystem-unfriendly characters
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    # Collapse whitespace to single space, strip
    name = " ".join(name.split()).strip()
    return name


def generate_1c_file(statement: Statement, output_dir: Path) -> Path:
    """Generate a 1CClientBankExchange file for a single statement.

    Output: output/{ClientName}-{PIB}/{account}_{stmt_number}_{date}.txt
    All accounts of the same company go into one directory.
    """
    # Directory: "CompanyName-PIB" or fallback to account number
    client = _safe_dirname(statement.client_name or "")
    pib = statement.client_pib or ""
    if client and pib:
        company_dir_name = f"{client}-{pib}"
    elif client:
        company_dir_name = client
    elif pib:
        company_dir_name = pib
    else:
        company_dir_name = _fmt_account(statement.account_number) or "unknown"

    company_dir = output_dir / company_dir_name
    company_dir.mkdir(parents=True, exist_ok=True)

    # Filename: account_stmtN_date.txt
    acct = _fmt_account(statement.account_number) or "unknown"
    stmt_num = statement.statement_number or str(statement.id)
    date_str = statement.statement_date.strftime("%Y%m%d") if statement.statement_date else "nodate"
    filename = f"{acct}_{stmt_num}_{date_str}.txt"
    output_path = company_dir / filename

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
    stmt_num = statement.statement_number or ""
    for tx in statement.transactions:
        is_debit = tx.debit is not None and tx.debit > 0
        amount = tx.debit if is_debit else tx.credit
        tx_date = tx.value_date or tx.booking_date

        # Номер: statement_number-row_number for unique identification
        doc_num = f"{stmt_num}-{tx.row_number}" if stmt_num else str(tx.row_number)
        lines.append("СекцияДокумент=Платёжное поручение")
        lines.append(f"Номер={doc_num}")
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

        purpose = _purpose_with_code(tx.payment_code, tx.purpose)
        lines.append(f"НазначениеПлатежа={_safe_text(purpose)}")
        lines.append("КонецДокумента")

    lines.append("КонецФайла")

    content = "\r\n".join(lines)
    output_path.write_text(content, encoding="windows-1251")

    return output_path
