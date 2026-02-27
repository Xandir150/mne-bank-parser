# Izvod — Montenegrin Bank Statement Parser for 1C

A web service that automatically parses PDF/HTML bank statements from Montenegrin banks and generates import files for 1C:Enterprise accounting software.

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
3. **Web interface** at `http://localhost:8000` shows parsed statements with inline editing
4. **Export to 1C** — generates `1CClientBankExchange` format files (UTF-8) for import into 1C:Бухгалтерия 3.0
5. **Originals are moved** to `data/processed/{bank_code}/` after parsing

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

# Wait ~60 seconds for automatic parsing, or check the web UI
```

## Architecture

```
izvod/
├── app/
│   ├── main.py          # FastAPI web application
│   ├── config.py        # Configuration (paths, bank codes)
│   ├── database.py      # SQLite + SQLAlchemy
│   ├── models.py        # Statement & Transaction ORM models
│   ├── worker.py        # Background directory scanner (APScheduler)
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
│   ├── processed/       # Originals moved here after parsing
│   ├── output/          # Generated 1C import files
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
- **Background Jobs:** APScheduler
- **Frontend:** Jinja2 + Bootstrap 5 + Tabulator.js (inline editing)
- **Container:** Docker (python:3.12-slim-bookworm)

## Web Interface

- **Dashboard** — list of all parsed statements with status badges (new/reviewed/exported/error)
- **Statement Detail** — editable header fields + transaction table with inline cell editing
- **Language switcher** — Russian (default) and Serbian
- **Auto-refresh** — dashboard updates every 30 seconds

## 1C Integration

The service generates files in `1CClientBankExchange` format (version 1.03, UTF-8 encoding).

### Setup in 1C:

1. Mount `data/output/` as a network drive or shared folder accessible from the 1C server
2. In 1C:Бухгалтерия 3.0, go to **Банк и касса → Обмен с банком → Настройка обмена**
3. Set the exchange directory to the mounted `data/output/` path
4. Click **Загрузить** to import parsed statements

### Output Format

```
1CClientBankExchange
ВерсияФормата=1.03
Кодировка=UTF-8
...
СекцияДокумент=Платёжное поручение
Номер=1
Дата=01.02.2026
Сумма=3.88
ПлательщикСчет=520000000004307069
Плательщик=FSTR DOO
...
КонецДокумента
КонецФайла
```

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
