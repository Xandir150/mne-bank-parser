from __future__ import annotations

import re
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING

import yaml

if TYPE_CHECKING:
    from app.models import Statement


# ---------------------------------------------------------------------------
# Config loading
# ---------------------------------------------------------------------------

_CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.yaml"
_config_cache = None


def _load_config() -> dict:
    global _config_cache
    if _config_cache is None:
        if _CONFIG_PATH.exists():
            with open(_CONFIG_PATH, encoding="utf-8") as f:
                _config_cache = yaml.safe_load(f) or {}
        else:
            _config_cache = {}
    return _config_cache


def _extract_rule(rule: dict, is_debit: bool) -> dict:
    """Extract fields from a config rule, resolving direction-dependent keys."""
    result = {}
    if is_debit:
        result["вид_операции"] = rule.get("вид_операции") or rule.get("вид_операции_дебет", "")
        result["статья_ддс"] = rule.get("статья_ддс") or rule.get("статья_ддс_дебет", "")
        result["счет_дебета"] = rule.get("счет_дебета") or rule.get("счет_дебета_дебет", "")
    else:
        result["вид_операции"] = rule.get("вид_операции") or rule.get("вид_операции_кредит", "")
        result["статья_ддс"] = rule.get("статья_ддс") or rule.get("статья_ддс_кредит", "")
        result["счет_дебета"] = rule.get("счет_дебета") or rule.get("счет_дебета_кредит", "")
    for key in ("вид_налога", "вид_обязательства", "статья_расходов"):
        if key in rule:
            result[key] = rule[key]
    return result


def _get_operation_info(counterparty_account: str, payment_code: str,
                        is_debit: bool, purpose: str = "",
                        counterparty: str = "") -> dict:
    """Determine operation type and related fields from config rules.

    Priority: 1) account pattern, 2) code + keyword, 3) payment code,
              4) purpose keyword, 5) counterparty keyword, 6) default.
    """
    cfg = _load_config()
    acct = _fmt_account(counterparty_account)

    # 1) Match by counterparty account pattern
    for rule in cfg.get("операции", []):
        pattern = rule.get("account_pattern", "")
        if pattern and acct and re.match(pattern, acct):
            return _extract_rule(rule, is_debit)

    # 2) Match by payment code + keyword in purpose
    purpose_lower = (purpose or "").lower()
    for rule in cfg.get("шифры_по_назначению", []):
        if payment_code == rule.get("шифра", "") and rule.get("слово", "").lower() in purpose_lower:
            return _extract_rule(rule, is_debit)

    # 3) Match by payment code
    code_rules = cfg.get("шифры", {})
    if payment_code and payment_code in code_rules:
        return _extract_rule(code_rules[payment_code], is_debit)

    # 4) Match by keyword in purpose only (no code required)
    for rule in cfg.get("по_назначению", []):
        if rule.get("слово", "").lower() in purpose_lower:
            return _extract_rule(rule, is_debit)

    # 5) Match by counterparty name keyword
    cp_lower = (counterparty or "").lower()
    for rule in cfg.get("по_контрагенту", []):
        if rule.get("слово", "").lower() in cp_lower:
            return _extract_rule(rule, is_debit)

    # 6) Default from config
    default = cfg.get("дефолт", {})
    if default:
        return _extract_rule(default, is_debit)

    # 5) Fallback from hardcoded cash flow items
    cash_flow = _CASH_FLOW_ITEM.get((payment_code, is_debit), "")
    if cash_flow:
        return {"статья_ддс": cash_flow}

    return {}


# ---------------------------------------------------------------------------
# Montenegrin payment codes (šifra plaćanja) → Russian descriptions for 1C
# ---------------------------------------------------------------------------

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

# Payment code → 1C "Статья движения денежных средств"
_CASH_FLOW_ITEM = {
    ("121", True):  "Оплата поставщикам (подрядчикам)",
    ("121", False): "Оплата от покупателей",
    ("122", True):  "Оплата поставщикам (подрядчикам)",
    ("122", False): "Оплата от покупателей",
    ("120", True):  "Оплата поставщикам (подрядчикам)",
    ("120", False): "Оплата от покупателей",
    ("151", True):  "Прочие расходы",
    ("151", False): "Прочие поступления",
    ("140", True):  "Прочие налоги и сборы",
    ("140", False): "Прочие налоги и сборы",
    ("139", True):  "Прочие налоги и сборы",
    ("139", False): "Прочие налоги и сборы",
    ("221", True):  "Расходы на услуги банков",
    ("221", False): "Расходы на услуги банков",
    ("400", True):  "Расходы на услуги банков",
    ("400", False): "Расходы на услуги банков",
    ("163", True):  "Прочие расходы",
    ("163", False): "Прочие поступления",
    ("170", True):  "Погашение кредитов и займов",
    ("170", False): "Получение кредитов и займов",
    ("171", True):  "Погашение кредитов и займов",
    ("171", False): "Получение кредитов и займов",
    ("126", True):  "Оплата поставщикам (подрядчикам)",
    ("126", False): "Оплата от покупателей",
    ("153", True):  "Прочие налоги и сборы",
    ("153", False): "Прочие налоги и сборы",
    ("165", True):  "Прочие расходы",
    ("165", False): "Розничная выручка",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

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

    Montenegrin accounts: BBB-NNNNNNNNNNNNN-CC (3-13-2 = 18 digits).
    See docs/account_rules.md for full specification.
    """
    if not acct:
        return ""
    # Strip IBAN prefix ME25
    s = acct.strip()
    if s.upper().startswith("ME"):
        s = re.sub(r"^ME\d{2}", "", s)
    # Remove all non-digit characters, then re-parse
    digits = re.sub(r"\D", "", s)
    if len(digits) == 18:
        return digits
    # Try dash-separated format: BBB-N...N-CC
    m = re.match(r"^(\d{3})\D+(\d+)\D+(\d{2})$", s)
    if m:
        bank, number, check = m.group(1), m.group(2), m.group(3)
        return bank + number.zfill(13) + check
    # Fallback: pad digits to 18 with zeros after bank code (first 3)
    if len(digits) > 3 and len(digits) < 18:
        return digits[:3] + digits[3:].zfill(15)
    return digits


def _fmt_pib(pib) -> str:
    """Format PIB (tax ID) for 1C: always 8 digits, zero-padded."""
    if not pib:
        return ""
    digits = re.sub(r"\D", "", str(pib))
    if not digits:
        return ""
    return digits.zfill(8)


def _fmt_amount(val) -> str:
    """Format decimal amount with 2 decimal places using dot separator."""
    if val is None:
        return "0.00"
    return f"{Decimal(val):.2f}"


def _safe_dirname(name: str) -> str:
    """Sanitize a string for use as a directory name."""
    if not name:
        return ""
    name = name.translate(_LATIN_MAP)
    name = re.sub(r'[<>:"/\\|?*]', '', name)
    name = " ".join(name.split()).strip()
    return name


# ---------------------------------------------------------------------------
# Main export
# ---------------------------------------------------------------------------

def generate_1c_file(statement: Statement, output_dir: Path) -> Path:
    """Generate a 1CClientBankExchange file for a single statement.

    Output: output/{ClientName}-{PIB}/{account}_{stmt_number}_{date}.txt
    """
    cfg = _load_config()
    счет_учета = cfg.get("счет_учета", "51")

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

    # --- Header ---
    lines.append("1CClientBankExchange")
    lines.append("ВерсияФормата=1.03")
    lines.append("Кодировка=Windows")
    lines.append("Отправитель=BankStatementParser")
    lines.append("Получатель=")
    lines.append(f"ДатаСоздания={now.strftime('%d.%m.%Y')}")
    lines.append(f"ВремяСоздания={now.strftime('%H:%M:%S')}")
    lines.append(f"ДатаНачала={_fmt_date(stmt_date)}")
    lines.append(f"ДатаКонца={_fmt_date(stmt_end)}")
    lines.append(f"РасчСчет={acct}")

    # --- Account section ---
    lines.append("СекцияРасчСчет")
    lines.append(f"ДатаНачала={_fmt_date(stmt_date)}")
    lines.append(f"ДатаКонца={_fmt_date(stmt_end)}")
    lines.append(f"РасчСчет={acct}")
    lines.append(f"НачальныйОстаток={_fmt_amount(statement.opening_balance)}")
    lines.append(f"КонечныйОстаток={_fmt_amount(statement.closing_balance)}")
    lines.append(f"ДебетОборот={_fmt_amount(statement.total_debit)}")
    lines.append(f"КредитОборот={_fmt_amount(statement.total_credit)}")
    lines.append("КонецРасчСчет")

    # --- Transactions ---
    stmt_num_str = statement.statement_number or ""
    for tx in statement.transactions:
        is_debit = tx.debit is not None and tx.debit > 0
        amount = tx.debit if is_debit else tx.credit
        tx_date = tx.value_date or tx.booking_date

        # Operation info from config
        op_info = _get_operation_info(
            tx.counterparty_account, tx.payment_code, is_debit, tx.purpose,
            tx.counterparty)

        # Номер документа = номер выписки
        doc_num = stmt_num_str or str(tx.row_number)

        lines.append("СекцияДокумент=Платёжное поручение")
        lines.append(f"Номер={doc_num}")
        lines.append(f"Дата={_fmt_date(tx_date)}")
        lines.append(f"Сумма={_fmt_amount(amount)}")

        # Номер и дата выписки
        lines.append(f"ДатаВыписки={_fmt_date(statement.statement_date)}")
        lines.append(f"НомерВыписки={statement.statement_number or ''}")

        # Счёт учёта (кредит — расчётный счёт)
        lines.append(f"СчетУчета={счет_учета}")

        # Счёт дебета (из правил конфига)
        счет_дебета = op_info.get("счет_дебета", "")
        if счет_дебета:
            lines.append(f"СчетДебета={счет_дебета}")

        # Вид операции
        вид_операции = op_info.get("вид_операции", "")
        if вид_операции:
            lines.append(f"ВидОперации={вид_операции}")

        # Плательщик / Получатель
        if is_debit:
            lines.append(f"ПлательщикСчет={acct}")
            lines.append(f"Плательщик={_safe_text(statement.client_name or '')}")
            lines.append(f"ПлательщикИНН={_fmt_pib(statement.client_pib)}")
            lines.append(f"ПолучательСчет={_fmt_account(tx.counterparty_account)}")
            lines.append(f"Получатель={_safe_text(tx.counterparty or '')}")
            lines.append("ПолучательИНН=")
        else:
            lines.append(f"ПлательщикСчет={_fmt_account(tx.counterparty_account)}")
            lines.append(f"Плательщик={_safe_text(tx.counterparty or '')}")
            lines.append("ПлательщикИНН=")
            lines.append(f"ПолучательСчет={acct}")
            lines.append(f"Получатель={_safe_text(statement.client_name or '')}")
            lines.append(f"ПолучательИНН={_fmt_pib(statement.client_pib)}")

        # Статья ДДС
        статья_ддс = op_info.get("статья_ддс", "")
        if статья_ддс:
            lines.append(f"СтатьяДвиженияДенежныхСредств={статья_ддс}")

        # Налоговые поля (только для "Уплата налога")
        if вид_операции == "Уплата налога":
            вид_налога = op_info.get("вид_налога", "")
            if вид_налога:
                lines.append(f"ВидНалога={вид_налога}")
            вид_обяз = op_info.get("вид_обязательства", "")
            if вид_обяз:
                lines.append(f"ВидОбязательства={вид_обяз}")
            статья_расх = op_info.get("статья_расходов", "")
            if статья_расх:
                lines.append(f"СтатьяРасходов={статья_расх}")

        # НазначениеПлатежа ПОСЛЕДНИМ — 1С читает его как многострочное
        # поле до КонецДокумента, всё после него попадает в текст назначения
        lines.append(f"НазначениеПлатежа={_safe_text(tx.purpose or '')}")

        lines.append("КонецДокумента")

    lines.append("КонецФайла")

    content = "\r\n".join(lines)
    output_path.write_text(content, encoding="windows-1251")

    return output_path
