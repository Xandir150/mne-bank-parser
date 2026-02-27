import logging
import shutil
from pathlib import Path

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger

from app.config import settings
from app.database import SessionLocal
from app.models import Statement, Transaction
from app.parsers import parse_file
from app.export_1c import generate_1c_file

logger = logging.getLogger(__name__)


def _setup_file_logging():
    """Configure logging to write to data/log/izvod.log."""
    log_dir = settings.data_dir / "log"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "izvod.log"

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    ))

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(file_handler)

scheduler = BackgroundScheduler()


def scan_directories():
    """Scan input directories for new bank statement files and process them."""
    db = SessionLocal()
    try:
        for bank_code in settings.bank_names:
            input_dir = settings.input_dir / bank_code
            if not input_dir.exists():
                continue

            extensions = settings.supported_extensions.get(bank_code, [".pdf"])
            for file_path in input_dir.iterdir():
                if not file_path.is_file():
                    continue
                if file_path.suffix.lower() not in extensions:
                    continue

                # Skip already processed files
                existing = (
                    db.query(Statement)
                    .filter(Statement.source_file == file_path.name, Statement.bank_code == bank_code)
                    .first()
                )
                if existing:
                    continue

                logger.info("Processing %s for bank %s", file_path.name, bank_code)
                try:
                    parsed = parse_file(file_path, bank_code)

                    stmt = Statement(
                        bank_code=parsed.bank_code,
                        bank_name=parsed.bank_name,
                        account_number=parsed.account_number,
                        iban=parsed.iban,
                        statement_number=parsed.statement_number,
                        statement_date=parsed.statement_date,
                        period_start=parsed.period_start,
                        period_end=parsed.period_end,
                        opening_balance=parsed.opening_balance,
                        closing_balance=parsed.closing_balance,
                        total_debit=parsed.total_debit,
                        total_credit=parsed.total_credit,
                        currency=parsed.currency,
                        client_name=parsed.client_name,
                        client_pib=parsed.client_pib,
                        source_file=file_path.name,
                        status="new",
                    )
                    db.add(stmt)
                    db.flush()

                    for pt in parsed.transactions:
                        tx = Transaction(
                            statement_id=stmt.id,
                            row_number=pt.row_number,
                            value_date=pt.value_date,
                            booking_date=pt.booking_date,
                            debit=pt.debit,
                            credit=pt.credit,
                            counterparty=pt.counterparty,
                            counterparty_account=pt.counterparty_account,
                            counterparty_bank=pt.counterparty_bank,
                            payment_code=pt.payment_code,
                            purpose=pt.purpose,
                            reference_debit=pt.reference_debit,
                            reference_credit=pt.reference_credit,
                            reclamation_data=pt.reclamation_data,
                            fee=pt.fee,
                        )
                        db.add(tx)

                    db.commit()

                    # Auto-export to 1C format
                    try:
                        export_path = generate_1c_file(stmt, settings.output_dir)
                        stmt.export_file = str(export_path)
                        stmt.status = "exported"
                        db.commit()
                        logger.info("Auto-exported -> %s", export_path)
                    except Exception as export_err:
                        logger.error("Auto-export failed for %s: %s", file_path.name, export_err)

                    # Move to processed
                    processed_dir = settings.processed_dir / bank_code
                    processed_dir.mkdir(parents=True, exist_ok=True)
                    dest = processed_dir / file_path.name
                    if dest.exists():
                        dest = processed_dir / f"{file_path.stem}_{stmt.id}{file_path.suffix}"
                    shutil.move(str(file_path), str(dest))

                    logger.info("Successfully processed %s -> Statement #%d", file_path.name, stmt.id)

                except Exception as e:
                    db.rollback()
                    logger.error("Failed to process %s: %s", file_path.name, e, exc_info=True)

                    # Save error record — file stays in input for retry
                    error_stmt = Statement(
                        bank_code=bank_code,
                        bank_name=settings.bank_names.get(bank_code, bank_code),
                        account_number="",
                        source_file=file_path.name,
                        status="error",
                        error_message=str(e),
                    )
                    db.add(error_stmt)
                    db.commit()
    finally:
        db.close()


def start_scheduler():
    """Start the background scheduler for directory scanning."""
    _setup_file_logging()
    scheduler.add_job(
        scan_directories,
        trigger=IntervalTrigger(seconds=settings.scan_interval),
        id="scan_directories",
        name="Scan input directories for new statements",
        replace_existing=True,
    )
    scheduler.start()
    logger.info("Background scheduler started (interval=%ds)", settings.scan_interval)


def stop_scheduler():
    """Stop the background scheduler."""
    if scheduler.running:
        scheduler.shutdown(wait=False)
        logger.info("Background scheduler stopped")
