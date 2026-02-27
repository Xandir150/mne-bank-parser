# Izvod — Montenegrin Bank Statement Parser for 1C

A web service that automatically parses PDF/HTML bank statements from Montenegrin banks and generates import files for 1C:Enterprise (Бухгалтерия 3.0).

Fully automatic pipeline: drop a PDF into the input folder → parsed → exported to 1CClientBankExchange format → ready for import in 1C.

## Supported Banks

| Code | Bank | Format | Notes |
|------|------|--------|-------|
| 520 | Hipotekarna Banka | PDF | Period decimal format |
| 530 | NLB Banka | PDF | CID-encoded fonts, custom character mapping |
| 535 | Prva Banka CG | PDF | RB-numbered rows with fee sub-lines |
| 540 | Erste Bank | HTML | Windows-1250 encoding |
| 560 | Universal Capital Bank | PDF | Multi-page statements |
| 565 | Lovćen Banka | PDF | Table-based extraction |
| 570 | Zapad Banka | PDF | Two sub-formats (daily + period) |
| 575 | Ziraat Bank Montenegro | PDF | RB-numbered with fee extraction |
| 580 | Adriatic Bank | PDF | English language |

## How It Works

1. **Drop PDF/HTML files** into `data/input/{bank_code}/` directories (e.g., `data/input/520/`)
2. **Background worker** scans directories every 60 seconds, parses new files automatically
3. **Auto-export** — immediately generates `1CClientBankExchange` files (Windows-1251) organized by account number
4. **Web interface** at `http://localhost:8000` shows parsed statements with inline editing (Russian/Serbian)
5. **Originals are moved** to `data/processed/{bank_code}/` after successful parsing
6. **Failed files stay** in `input/` for retry, errors are logged to `data/log/izvod.log`

## Quick Start

```bash
# Clone and run
git clone https://github.com/Xandir150/mne-bank-parser.git
cd mne-bank-parser
docker compose up -d --build

# Open web UI
open http://localhost:8000

# Drop a bank statement PDF into the appropriate directory
cp your_statement.pdf data/input/520/

# Wait ~60 seconds — file is parsed, exported, and ready for 1C
```

## 1C Integration

The service generates files in `1CClientBankExchange` format (version 1.03, Windows-1251 encoding).

### Output Structure

Each statement is exported to its own file, organized by account number:

```
data/output/
├── 520000000004307069/
│   └── 20260201_1.txt      ← one file per statement
├── 560000000000290342/
│   ├── 20260202_5.txt
│   └── 20260215_12.txt     ← new statements get new files
└── ...
```

Account numbers are normalized to 18-digit format without dashes (e.g., `535-22023-67` → `535000000002202367`).

### Setup in 1C (per bank account)

In 1C:Бухгалтерия 3.0, go to **Банк и касса → Обмен с банком → Настройка обмена с клиентом банка**:

| Field | Value |
|-------|-------|
| Обслуживаемый банковский счет | Select the account |
| Файл выгрузки в банк | _(leave empty)_ |
| Файл загрузки из банка | `\\server\izvod\data\output\{account}\{file}.txt` |
| Кодировка | **Windows** |

**Recommended checkboxes:**
- **Автоматическое создание ненайденных элементов** — yes (auto-creates new counterparties)
- **Перед загрузкой показывать форму "Обмен с банком"** — yes (review before import)

To import: click **Загрузить** in bank statements, browse to the account folder, select the file.

### Output File Format

```
1CClientBankExchange
ВерсияФормата=1.03
Кодировка=Windows
Отправитель=BankStatementParser
...
РасчСчет=520000000004307069
СекцияРасчСчет
НачальныйОстаток=1069.94
КонечныйОстаток=1066.06
ДебетОборот=3.88
КредитОборот=0.00
КонецРасчСчет
СекцияДокумент=Платёжное поручение
Номер=1
Дата=01.02.2026
Сумма=3.88
ПлательщикСчет=520000000004307069
Плательщик=FSTR DOO
ПлательщикИНН=03424804
ПолучательСчет=520000000004307069
Получатель=...
НазначениеПлатежа=Naplata naknada
КонецДокумента
КонецФайла
```

## Architecture

```
izvod/
├── app/
│   ├── main.py          # FastAPI web application
│   ├── config.py        # Configuration (paths, bank codes)
│   ├── database.py      # SQLite + SQLAlchemy
│   ├── models.py        # Statement & Transaction ORM models
│   ├── worker.py        # Background scanner + auto-export (APScheduler)
│   ├── export_1c.py     # 1CClientBankExchange file generator
│   ├── i18n.py          # Russian & Serbian translations
│   ├── parsers/
│   │   ├── base.py      # Abstract parser + ParsedStatement/ParsedTransaction
│   │   ├── hipotekarna.py   # 520
│   │   ├── nlb.py           # 530
│   │   ├── prva.py          # 535
│   │   ├── erste.py         # 540 (HTML)
│   │   ├── ucb.py           # 560
│   │   ├── lovcen.py        # 565
│   │   ├── zapad.py         # 570
│   │   ├── ziraat.py        # 575
│   │   └── adriatic.py      # 580
│   ├── templates/       # Jinja2 templates (Bootstrap 5 + Tabulator.js)
│   └── static/          # CSS
├── data/
│   ├── input/           # Drop bank statements here (by bank code)
│   ├── processed/       # Originals moved here after successful parsing
│   ├── output/          # Generated 1C files (by account number)
│   ├── log/             # Application log (izvod.log)
│   └── db/              # SQLite database
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```

## Tech Stack

- **Backend:** Python 3.12 + FastAPI
- **PDF Parsing:** pdfplumber
- **HTML Parsing:** BeautifulSoup4 (for Erste Bank)
- **Database:** SQLite + SQLAlchemy
- **Background Jobs:** APScheduler (60s interval)
- **Frontend:** Jinja2 + Bootstrap 5 + Tabulator.js (inline editing)
- **Container:** Docker (python:3.12-slim-bookworm)

## Web Interface

- **Dashboard** — list of all parsed statements with status badges (new/reviewed/exported/error)
- **Statement Detail** — editable header fields + transaction table with inline cell editing
- **Language switcher** — Russian (default) and Serbian
- **Auto-refresh** — dashboard updates every 30 seconds

## Configuration

Environment variables (set in `docker-compose.yml`):

| Variable | Default | Description |
|----------|---------|-------------|
| `IZVOD_SCAN_INTERVAL` | `60` | Directory scan interval in seconds |

## Extracted Fields

### Statement Header
Account number, IBAN, statement number/date, period start/end, opening/closing balance, total debit/credit, currency, client name, PIB (tax ID)

### Transaction
Row number, value date, booking date, debit, credit, counterparty name/account/bank, payment code, purpose, debit/credit references, reclamation data, fee

## Error Handling

- **Parse errors**: file stays in `input/`, statement with status `error` shown in web UI, full traceback in `data/log/izvod.log`
- **Export errors**: statement is parsed and saved, but export file is not generated; can be re-exported from the web UI
- **Serbian characters** (š, č, ž, ć, đ): transliterated to Latin equivalents for Windows-1251 compatibility

## Adding a New Bank Parser

1. Create `app/parsers/newbank.py`
2. Extend `BankParser`, set `bank_code` and `bank_name`
3. Implement `parse(file_path) -> ParsedStatement`
4. Add `@register_parser` decorator
5. Add bank code to `settings.bank_names` in `app/config.py`

```python
from app.parsers import register_parser
from app.parsers.base import BankParser, ParsedStatement

@register_parser
class NewBankParser(BankParser):
    bank_code = "999"
    bank_name = "New Bank"

    def parse(self, file_path):
        stmt = ParsedStatement(bank_code=self.bank_code, bank_name=self.bank_name)
        # ... parse logic ...
        return stmt
```

## License

MIT
