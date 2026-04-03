"""Export bank statements to EnterpriseData XML format for 1C:Бухгалтерия 3.0.

Generates XML with documents:
- СписаниеСРасчетногоСчета (debit / money out)
- ПоступлениеНаРасчетныйСчет (credit / money in)

All catalog references use natural keys (name, account number, INN)
so 1C can match them to existing entries.
"""
import re
import uuid
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING
from xml.etree.ElementTree import Element, SubElement, tostring
from xml.dom.minidom import parseString

import yaml

if TYPE_CHECKING:
    from app.models import Statement

# ---------------------------------------------------------------------------
# Config (shared with export_1c.py)
# ---------------------------------------------------------------------------

_CONFIG_PATH = Path(__file__).resolve().parent.parent / "config.yaml"
_config_cache = None
_config_mtime = 0.0


def _load_config() -> dict:
    global _config_cache, _config_mtime
    try:
        mtime = _CONFIG_PATH.stat().st_mtime
    except OSError:
        mtime = 0.0
    if _config_cache is None or mtime != _config_mtime:
        if _CONFIG_PATH.exists():
            with open(_CONFIG_PATH, encoding="utf-8") as f:
                _config_cache = yaml.safe_load(f) or {}
        else:
            _config_cache = {}
        _config_mtime = mtime
    return _config_cache


def _fmt_account(acct) -> str:
    if not acct:
        return ""
    s = acct.strip()
    if s.upper().startswith("ME"):
        s = re.sub(r"^ME\d{2}", "", s)
    digits = re.sub(r"\D", "", s)
    if len(digits) == 18:
        return digits
    m = re.match(r"^(\d{3})\D+(\d+)\D+(\d{2})$", s)
    if m:
        bank, number, check = m.group(1), m.group(2), m.group(3)
        return bank + number.zfill(13) + check
    if 3 < len(digits) < 18:
        return digits[:3] + digits[3:].zfill(15)
    return digits


def _fmt_pib(pib) -> str:
    if not pib:
        return ""
    digits = re.sub(r"\D", "", str(pib))
    return digits.zfill(8) if digits else ""


def _new_guid() -> str:
    return str(uuid.uuid4())


# ---------------------------------------------------------------------------
# Operation type mapping (same logic as export_1c.py)
# ---------------------------------------------------------------------------

# ВидОперации enum values for 1C (internal names)
_OP_MAP = {
    "Уплата налога": "ПеречислениеНалога",
    "Комиссия банка": "ПрочееСписание",
    "Оплата поставщику": "ОплатаПоставщику",
    "Оплата от покупателя": "ОплатаОтПокупателя",
    "Прочее списание": "ПрочееСписание",
    "Прочее поступление": "ПрочееПоступление",
}


def _extract_rule(rule: dict, is_debit: bool) -> dict:
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
    cfg = _load_config()
    acct = _fmt_account(counterparty_account)

    for rule in cfg.get("операции", []):
        pattern = rule.get("account_pattern", "")
        if pattern and acct and re.match(pattern, acct):
            return _extract_rule(rule, is_debit)

    purpose_lower = (purpose or "").lower()
    for rule in cfg.get("шифры_по_назначению", []):
        if payment_code == rule.get("шифра", "") and rule.get("слово", "").lower() in purpose_lower:
            return _extract_rule(rule, is_debit)

    code_rules = cfg.get("шифры", {})
    if payment_code and payment_code in code_rules:
        return _extract_rule(code_rules[payment_code], is_debit)

    for rule in cfg.get("по_назначению", []):
        if rule.get("слово", "").lower() in purpose_lower:
            return _extract_rule(rule, is_debit)

    cp_lower = (counterparty or "").lower()
    for rule in cfg.get("по_контрагенту", []):
        if rule.get("слово", "").lower() in cp_lower:
            return _extract_rule(rule, is_debit)

    default = cfg.get("дефолт", {})
    if default:
        return _extract_rule(default, is_debit)

    return {}


# ---------------------------------------------------------------------------
# Bank code → bank name mapping
# ---------------------------------------------------------------------------

_BANK_NAMES = {
    "510": "Centralna Banka Crne Gore",
    "520": "Hipotekarna Banka",
    "530": "NLB Banka",
    "535": "Prva Banka",
    "540": "Erste Bank",
    "560": "Universal Capital Bank",
    "565": "Lovcen Banka",
    "570": "Zapad Banka",
    "575": "Ziraat Bank",
    "580": "Adriatic Bank",
    "820": "Poreska uprava Crne Gore",
    "907": "Lovcen Banka (interni)",
}


def _bank_name_for_account(acct: str) -> str:
    acct = _fmt_account(acct)
    if len(acct) >= 3:
        return _BANK_NAMES.get(acct[:3], "")
    return ""


# ---------------------------------------------------------------------------
# XML builder helpers
# ---------------------------------------------------------------------------

NS = "http://v8.1c.ru/edi/edi_stnd/EnterpriseData/1.12"
MSG_NS = "http://www.1c.ru/SSL/Exchange/Message"


def _add_text(parent: Element, tag: str, text: str) -> Element:
    el = SubElement(parent, tag)
    el.text = text
    return el


def _add_ref(parent: Element, tag: str, name: str = "",
             inn: str = "", account: str = "") -> Element:
    """Add a catalog reference element with natural keys."""
    el = SubElement(parent, tag)
    _add_text(el, "Наименование", name)
    if inn:
        _add_text(el, "ИНН", inn)
    if account:
        _add_text(el, "НомерСчета", account)
    return el


def _add_bank_account(parent: Element, tag: str, account_number: str,
                      bank_name: str = "") -> Element:
    el = SubElement(parent, tag)
    _add_text(el, "НомерСчета", _fmt_account(account_number))
    if bank_name:
        bank_el = SubElement(el, "Банк")
        _add_text(bank_el, "Наименование", bank_name)
    return el


# ---------------------------------------------------------------------------
# Main export
# ---------------------------------------------------------------------------

def generate_xml_file(statement, output_dir: Path) -> Path:
    """Generate EnterpriseData XML file for a single bank statement."""
    cfg = _load_config()

    # Sanitize names for directory
    client_name = statement.client_name or ""
    pib = _fmt_pib(statement.client_pib)
    acct = _fmt_account(statement.account_number) or "unknown"

    if client_name and pib:
        dirname = f"{client_name}-{pib}"
    elif client_name:
        dirname = client_name
    else:
        dirname = acct
    # Clean dirname
    dirname = re.sub(r'[<>:"/\\|?*]', '', dirname).strip()

    company_dir = output_dir / dirname
    company_dir.mkdir(parents=True, exist_ok=True)

    stmt_num = statement.statement_number or "0"
    date_str = statement.statement_date.strftime("%Y%m%d") if statement.statement_date else "nodate"
    filename = f"{acct}_{stmt_num}_{date_str}.xml"
    output_path = company_dir / filename

    stmt_date = statement.statement_date
    org_bank_name = _bank_name_for_account(acct)

    # Build XML
    root = Element("Message")
    root.set("xmlns", MSG_NS)
    root.set("xmlns:msg", MSG_NS)
    root.set("xmlns:ns1", NS)

    # Header
    header = SubElement(root, "msg:Header")
    _add_text(header, "msg:Format", NS)
    _add_text(header, "msg:CreationDate",
              datetime.now().strftime("%Y-%m-%dT%H:%M:%S"))

    conf = SubElement(header, "msg:Confirmation")
    _add_text(conf, "msg:ExchangePlan",
              "СинхронизацияДанныхЧерезУниверсальныйФормат")
    _add_text(conf, "msg:To", "БП")
    _add_text(conf, "msg:From", "BankParser")
    _add_text(conf, "msg:MessageNo", "1")
    _add_text(conf, "msg:ReceivedNo", "0")

    _add_text(header, "msg:AvailableVersion", "1.12")

    body = SubElement(root, "msg:Body")

    # Generate one document per transaction
    for tx in statement.transactions:
        is_debit = tx.debit is not None and tx.debit > 0
        amount = tx.debit if is_debit else tx.credit
        tx_date = tx.value_date or tx.booking_date

        op_info = _get_operation_info(
            tx.counterparty_account, tx.payment_code, is_debit, tx.purpose,
            tx.counterparty)
        вид_операции_ru = op_info.get("вид_операции", "")
        вид_операции = _OP_MAP.get(вид_операции_ru, "ПрочееСписание" if is_debit else "ПрочееПоступление")

        # Document type depends on direction
        if is_debit:
            doc_tag = f"ns1:Документ.СписаниеСРасчетногоСчета"
        else:
            doc_tag = f"ns1:Документ.ПоступлениеНаРасчетныйСчет"

        doc = SubElement(body, doc_tag)

        # КлючевыеСвойства
        key_props = SubElement(doc, "КлючевыеСвойства")
        _add_text(key_props, "Ссылка", _new_guid())
        _add_text(key_props, "Дата",
                  tx_date.strftime("%Y-%m-%dT00:00:00") if tx_date else "")
        _add_text(key_props, "Номер", stmt_num)

        # Организация
        org = SubElement(key_props, "Организация")
        _add_text(org, "Наименование", client_name)
        if pib:
            _add_text(org, "ИНН", pib)

        # ВидОперации
        _add_text(doc, "ВидОперации", вид_операции)

        # Сумма
        _add_text(doc, "СуммаДокумента",
                  f"{Decimal(amount):.2f}" if amount else "0.00")

        # Валюта
        _add_ref(doc, "Валюта", name="EUR")

        # Контрагент
        cp = SubElement(doc, "Контрагент")
        _add_text(cp, "Наименование", tx.counterparty or "")

        # Банковский счёт организации
        _add_bank_account(doc, "БанковскийСчетОрганизации",
                          acct, org_bank_name)

        # Банковский счёт контрагента
        cp_bank = _bank_name_for_account(tx.counterparty_account or "")
        _add_bank_account(doc, "БанковскийСчетКонтрагента",
                          tx.counterparty_account or "", cp_bank)

        # Подразделение
        if client_name:
            _add_ref(doc, "Подразделение", name=client_name)

        # СтатьяДвиженияДенежныхСредств
        статья_ддс = op_info.get("статья_ддс", "")
        if статья_ддс:
            _add_ref(doc, "СтатьяДвиженияДенежныхСредств", name=статья_ддс)

        # Налоговые поля (для УплатаНалога / ПеречислениеНалога)
        if вид_операции == "ПеречислениеНалога":
            вид_налога = op_info.get("вид_налога", "")
            if вид_налога:
                _add_ref(doc, "ВидНалога", name=вид_налога)
            вид_обяз = op_info.get("вид_обязательства", "")
            if вид_обяз:
                _add_text(doc, "ВидОбязательства", вид_обяз)
            статья_расх = op_info.get("статья_расходов", "")
            if статья_расх:
                _add_ref(doc, "СтатьяРасходов", name=статья_расх)
            счет_дебета = op_info.get("счет_дебета", "")
            if счет_дебета:
                _add_text(doc, "СчетУчета", счет_дебета)

        # НазначениеПлатежа
        _add_text(doc, "НазначениеПлатежа", tx.purpose or "")

        # НомерВыписки, ДатаВыписки
        _add_text(doc, "НомерВыписки", stmt_num)
        if stmt_date:
            _add_text(doc, "ДатаВыписки",
                      stmt_date.strftime("%Y-%m-%d"))

    # Pretty-print XML
    raw = tostring(root, encoding="unicode")
    pretty = parseString(raw).toprettyxml(indent="  ", encoding="UTF-8")

    output_path.write_bytes(pretty)
    return output_path
