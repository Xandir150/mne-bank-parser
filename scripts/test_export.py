#!/usr/bin/env python3
"""Quick local test: parse a PDF and generate 1C export file."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.parsers.lovcen import LovcenParser
from app.parsers.base import ParsedStatement
from app.export_1c import generate_1c_file
from app.export_xml import generate_xml_file

# Stub Statement object that mimics the ORM model
class FakeStatement:
    def __init__(self, parsed: ParsedStatement):
        self.id = 1
        self.account_number = parsed.account_number
        self.client_name = parsed.client_name
        self.client_pib = parsed.client_pib
        self.statement_number = parsed.statement_number
        self.statement_date = parsed.statement_date
        self.period_start = parsed.period_start or parsed.statement_date
        self.period_end = parsed.period_end or parsed.statement_date
        self.opening_balance = parsed.opening_balance
        self.closing_balance = parsed.closing_balance
        self.total_debit = parsed.total_debit
        self.total_credit = parsed.total_credit
        self.transactions = []
        for tx in parsed.transactions:
            self.transactions.append(FakeTransaction(tx))

class FakeTransaction:
    def __init__(self, tx):
        self.row_number = tx.row_number
        self.value_date = tx.value_date
        self.booking_date = tx.booking_date
        self.debit = tx.debit
        self.credit = tx.credit
        self.counterparty = tx.counterparty
        self.counterparty_account = tx.counterparty_account
        self.counterparty_bank = tx.counterparty_bank
        self.payment_code = tx.payment_code
        self.purpose = tx.purpose
        self.reference_debit = tx.reference_debit
        self.reference_credit = tx.reference_credit
        self.reclamation_data = tx.reclamation_data


def main():
    pdf = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(
        "data/input/Izvod prometa - POS_565000000002231087_120326.pdf")

    if not pdf.exists():
        print(f"File not found: {pdf}")
        return

    print(f"Parsing: {pdf.name}")
    parser = LovcenParser()
    parsed = parser.parse(pdf)

    print(f"  Account: {parsed.account_number}")
    print(f"  Client: {parsed.client_name} PIB={parsed.client_pib}")
    print(f"  Statement #{parsed.statement_number} date={parsed.statement_date}")
    print(f"  Transactions: {len(parsed.transactions)}")

    stmt = FakeStatement(parsed)
    output_dir = Path("data/test_output")
    output_dir.mkdir(parents=True, exist_ok=True)

    # TXT (1CClientBankExchange)
    result_txt = generate_1c_file(stmt, output_dir)
    print(f"\nTXT: {result_txt}")

    # XML (EnterpriseData)
    result_xml = generate_xml_file(stmt, output_dir)
    print(f"XML: {result_xml}")

    # Print XML
    content = result_xml.read_text(encoding="utf-8")
    print("\n" + "=" * 60)
    print(content)


if __name__ == "__main__":
    main()
