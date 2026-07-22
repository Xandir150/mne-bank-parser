import logging
import shutil
from decimal import Decimal
from pathlib import Path
from typing import Optional

from apscheduler.schedulers.background import BackgroundScheduler
from apscheduler.triggers.interval import IntervalTrigger

from app.config import settings
from app.database import SessionLocal
from app.models import Statement, Transaction
from app.parsers import parse_file_multi, detect_bank_code
from app.export_1c import generate_1c_file

logger = logging.getLogger(__name__)

ALL_EXTENSIONS = {".pdf", ".htm", ".html"}


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


def _find_duplicate_statement(db, parsed, file_path: Path, bank_code: str) -> Optional[Statement]:
    """Find an already-imported Statement that IS this same document (safe to skip).

    Identity is (bank_code, account_number, statement_number, statement_date) —
    the bank's own numbering plus its date — not the source filename. Some
    banks reuse filenames across different accounts, and some (see bundled
    "whole period" exports) reuse a statement number across different dates
    for the SAME account, so number alone isn't a safe key. Falls back to
    source_file when account/statement number weren't extracted (parser gap
    on some format).
    """
    if parsed.account_number and parsed.statement_number:
        q = db.query(Statement).filter(
            Statement.bank_code == bank_code,
            Statement.account_number == parsed.account_number,
            Statement.statement_number == parsed.statement_number,
            Statement.status != "error",
        )
        if parsed.statement_date:
            q = q.filter(Statement.statement_date == parsed.statement_date)
        return q.first()
    return (
        db.query(Statement)
        .filter(
            Statement.source_file == file_path.name,
            Statement.bank_code == bank_code,
            Statement.status != "error",
        )
        .first()
    )


def _insert_statement(db, parsed, file_path: Path) -> Statement:
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

    try:
        export_path = generate_1c_file(stmt, settings.output_dir)
        stmt.export_file = str(export_path)
        stmt.status = "exported"
        db.commit()
        logger.info("Auto-exported -> %s", export_path)
    except Exception as export_err:
        logger.error(
            "Auto-export failed for %s (izvod %s): %s",
            file_path.name, parsed.statement_number, export_err,
        )

    return stmt


def _validate_balance(parsed) -> Optional[str]:
    """Reconcile the statement's printed control figures against the parsed
    transactions. The invariant `opening + credits - debits == closing` holds
    for every real bank statement, so a mismatch means the parser mangled the
    file (misaligned columns, lost rows, doubled rows) — such data must never
    reach 1C. Returns an error description, or None when consistent or when
    the parser didn't extract the control figures (nothing to check against).
    """
    if parsed.opening_balance is None or parsed.closing_balance is None:
        return None

    sum_debit = sum((t.debit or Decimal("0")) for t in parsed.transactions)
    sum_credit = sum((t.credit or Decimal("0")) for t in parsed.transactions)
    expected_closing = parsed.opening_balance + sum_credit - sum_debit
    tolerance = Decimal("0.01")

    if abs(expected_closing - parsed.closing_balance) > tolerance:
        return (
            f"balance mismatch: opening {parsed.opening_balance} + credits {sum_credit} "
            f"- debits {sum_debit} = {expected_closing}, but statement says closing "
            f"{parsed.closing_balance}"
        )

    # Secondary check when the statement also prints turnover totals
    if parsed.total_debit is not None and abs(sum_debit - parsed.total_debit) > tolerance:
        return (
            f"debit turnover mismatch: transactions sum to {sum_debit}, "
            f"statement says {parsed.total_debit}"
        )
    if parsed.total_credit is not None and abs(sum_credit - parsed.total_credit) > tolerance:
        return (
            f"credit turnover mismatch: transactions sum to {sum_credit}, "
            f"statement says {parsed.total_credit}"
        )
    return None


def _process_file(db, file_path: Path, bank_code: str) -> None:
    """Process a bank statement file. A file may bundle multiple statements

    (e.g. a bank's "whole period" export containing several complete IZVOD
    documents back to back) — each embedded statement is handled on its own:
    deduped, inserted, and exported independently.
    """
    # Clear out a previous failed attempt for this exact file so it can be retried
    existing_error = (
        db.query(Statement)
        .filter(
            Statement.source_file == file_path.name,
            Statement.bank_code == bank_code,
            Statement.status == "error",
        )
        .first()
    )
    if existing_error:
        db.delete(existing_error)
        db.commit()
        logger.info("Retrying previously failed file %s", file_path.name)

    logger.info("Processing %s for bank %s", file_path.name, bank_code)
    try:
        parsed_list = parse_file_multi(file_path, bank_code)
        is_bundle = len(parsed_list) > 1

        usable = [p for p in parsed_list if p.transactions]
        if not usable:
            raise ValueError(
                f"Parser returned 0 transactions for {file_path.name} — "
                "possible format change or extraction failure"
            )
        if len(usable) < len(parsed_list):
            logger.warning(
                "%s: %d of %d embedded statement(s) had no transactions and were skipped",
                file_path.name, len(parsed_list) - len(usable), len(parsed_list),
            )

        # Reconcile control figures before anything is written: one garbled
        # embedded statement discredits the whole file, so fail it loudly
        # (the except-branch records the error and moves the file to failed/).
        for parsed in usable:
            balance_err = _validate_balance(parsed)
            if balance_err:
                raise ValueError(
                    f"izvod {parsed.statement_number} (account {parsed.account_number}, "
                    f"date {parsed.statement_date}) in {file_path.name}: {balance_err}"
                )

        created_ids = []
        skipped_dup_ids = []
        held_for_review = []

        for parsed in usable:
            dup = _find_duplicate_statement(db, parsed, file_path, bank_code)
            if dup:
                logger.warning(
                    "Skipping izvod %s (account %s, date %s) in %s — already imported as Statement #%d",
                    parsed.statement_number, parsed.account_number, parsed.statement_date,
                    file_path.name, dup.id,
                )
                skipped_dup_ids.append(dup.id)
                continue

            # A bundled multi-statement file mixing zero-balance entries with
            # real ones is a red flag: on the one bundle we've seen this in,
            # the zero-balance "statement" turned out to be a documentation
            # duplicate of a transaction already posted in a real (non-zero)
            # sibling statement in the SAME bundle — importing both would
            # double-book it. Hold these back for manual review rather than
            # guess; a lone zero-balance file (not part of a bundle) is the
            # normal, already-supported case and is unaffected by this check.
            if (
                is_bundle
                and parsed.opening_balance == Decimal("0")
                and parsed.closing_balance == Decimal("0")
            ):
                logger.warning(
                    "Holding izvod %s (account %s, date %s) in %s for manual review — "
                    "zero-balance statement inside a multi-statement bundle, possible "
                    "duplicate of a real sibling statement",
                    parsed.statement_number, parsed.account_number, parsed.statement_date,
                    file_path.name,
                )
                held_for_review.append(parsed)
                continue

            stmt = _insert_statement(db, parsed, file_path)
            created_ids.append(stmt.id)
            logger.info(
                "Parsed izvod %s (account %s, date %s) from %s -> Statement #%d",
                parsed.statement_number, parsed.account_number, parsed.statement_date,
                file_path.name, stmt.id,
            )

        # Decide where the physical file ends up. Prefer normal `processed/`
        # whenever anything new was imported; `duplicates/` only if the
        # entire file turned out to already be known; `needs_review/` if
        # anything was held back for a human to look at.
        if held_for_review:
            dest_dir = settings.processed_dir / bank_code / "needs_review"
        elif created_ids:
            dest_dir = settings.processed_dir / bank_code
        else:
            dest_dir = settings.processed_dir / bank_code / "duplicates"
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / file_path.name
        if dest.exists():
            tag = created_ids[0] if created_ids else (skipped_dup_ids or ["x"])[0]
            dest = dest_dir / f"{file_path.stem}_{tag}{file_path.suffix}"
        shutil.move(str(file_path), str(dest))

        logger.info(
            "Successfully processed %s -> %d new statement(s) %s, %d duplicate(s) skipped, "
            "%d held for review",
            file_path.name, len(created_ids), created_ids, len(skipped_dup_ids), len(held_for_review),
        )

    except Exception as e:
        db.rollback()
        logger.error("Failed to process %s: %s", file_path.name, e, exc_info=True)

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

        # Move the file out of input/ so it stops being re-parsed (and re-failing)
        # on every scan cycle. It lands in processed/<bank>/failed/ for a human to
        # review; if still present in input/ (e.g. moved away in the meantime by
        # another process), that's fine — nothing left to do.
        if file_path.exists():
            try:
                dest_dir = settings.processed_dir / bank_code / "failed"
                dest_dir.mkdir(parents=True, exist_ok=True)
                dest = dest_dir / file_path.name
                if dest.exists():
                    dest = dest_dir / f"{file_path.stem}_{error_stmt.id}{file_path.suffix}"
                shutil.move(str(file_path), str(dest))
                logger.info("Moved failed file %s -> %s", file_path.name, dest)
            except Exception as move_err:
                logger.error(
                    "Could not move failed file %s to processed/%s/failed/: %s "
                    "— it will be retried next scan cycle",
                    file_path.name, bank_code, move_err,
                )


def scan_directories():
    """Scan input directories for new bank statement files and process them."""
    db = SessionLocal()
    try:
        # 1. Scan per-bank subdirectories (existing behavior)
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
                _process_file(db, file_path, bank_code)

        # 2. Scan root input/ for files with auto-detection
        for file_path in settings.input_dir.iterdir():
            if not file_path.is_file():
                continue
            if file_path.suffix.lower() not in ALL_EXTENSIONS:
                continue

            bank_code = detect_bank_code(file_path)
            if not bank_code:
                logger.warning(
                    "Could not auto-detect bank for %s — move to a bank folder or check format",
                    file_path.name,
                )
                continue

            logger.info("Auto-detected bank %s for %s", bank_code, file_path.name)
            _process_file(db, file_path, bank_code)
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
