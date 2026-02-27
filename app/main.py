from pathlib import Path

from fastapi import FastAPI, Request, Depends, HTTPException, Query
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session

from app.config import settings
from app.database import init_db, get_db
from app.models import Statement, Transaction
from app.worker import start_scheduler, stop_scheduler
from app.export_1c import generate_1c_file_multi
from app.i18n import get_translations, DEFAULT_LANG

app = FastAPI(title="Izvod - Bank Statement Parser")

templates = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))
app.mount("/static", StaticFiles(directory=str(Path(__file__).parent / "static")), name="static")


@app.on_event("startup")
def on_startup():
    init_db()
    for code in settings.bank_names:
        (settings.input_dir / code).mkdir(parents=True, exist_ok=True)
        (settings.processed_dir / code).mkdir(parents=True, exist_ok=True)
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    settings.db_path.parent.mkdir(parents=True, exist_ok=True)
    start_scheduler()


@app.on_event("shutdown")
def on_shutdown():
    stop_scheduler()


def _get_lang(request: Request, lang: str = None) -> str:
    """Determine language from query param or cookie."""
    if lang:
        return lang
    return request.cookies.get("lang", DEFAULT_LANG)


# ─── HTML Pages ───


@app.get("/", response_class=HTMLResponse)
def index(request: Request, lang: str = Query(None), db: Session = Depends(get_db)):
    lang = _get_lang(request, lang)
    t = get_translations(lang)
    statements = (
        db.query(Statement)
        .order_by(Statement.created_at.desc())
        .all()
    )
    response = templates.TemplateResponse("index.html", {
        "request": request,
        "statements": statements,
        "t": t,
        "lang": lang,
    })
    response.set_cookie("lang", lang, max_age=365 * 24 * 3600)
    return response


@app.get("/statement/{statement_id}", response_class=HTMLResponse)
def statement_page(statement_id: int, request: Request, lang: str = Query(None), db: Session = Depends(get_db)):
    lang = _get_lang(request, lang)
    t = get_translations(lang)
    stmt = db.query(Statement).filter(Statement.id == statement_id).first()
    if not stmt:
        raise HTTPException(status_code=404, detail="Statement not found")
    response = templates.TemplateResponse("statement.html", {
        "request": request,
        "statement": stmt,
        "transactions": stmt.transactions,
        "t": t,
        "lang": lang,
    })
    response.set_cookie("lang", lang, max_age=365 * 24 * 3600)
    return response


# ─── API ───


@app.get("/api/statements")
def api_statements(db: Session = Depends(get_db)):
    statements = db.query(Statement).order_by(Statement.created_at.desc()).all()
    return [
        {
            "id": s.id,
            "bank_code": s.bank_code,
            "bank_name": s.bank_name,
            "account_number": s.account_number,
            "statement_date": s.statement_date.strftime("%d.%m.%Y") if s.statement_date else None,
            "client_name": s.client_name,
            "currency": s.currency,
            "opening_balance": str(s.opening_balance) if s.opening_balance is not None else None,
            "closing_balance": str(s.closing_balance) if s.closing_balance is not None else None,
            "total_debit": str(s.total_debit) if s.total_debit is not None else None,
            "total_credit": str(s.total_credit) if s.total_credit is not None else None,
            "status": s.status,
            "source_file": s.source_file,
            "tx_count": len(s.transactions),
            "created_at": s.created_at.isoformat(),
        }
        for s in statements
    ]


@app.get("/api/statement/{statement_id}")
def api_statement(statement_id: int, db: Session = Depends(get_db)):
    stmt = db.query(Statement).filter(Statement.id == statement_id).first()
    if not stmt:
        raise HTTPException(status_code=404, detail="Statement not found")
    return {
        "id": stmt.id,
        "bank_code": stmt.bank_code,
        "bank_name": stmt.bank_name,
        "account_number": stmt.account_number,
        "iban": stmt.iban,
        "statement_number": stmt.statement_number,
        "statement_date": stmt.statement_date.isoformat() if stmt.statement_date else None,
        "period_start": stmt.period_start.isoformat() if stmt.period_start else None,
        "period_end": stmt.period_end.isoformat() if stmt.period_end else None,
        "opening_balance": str(stmt.opening_balance) if stmt.opening_balance is not None else None,
        "closing_balance": str(stmt.closing_balance) if stmt.closing_balance is not None else None,
        "total_debit": str(stmt.total_debit) if stmt.total_debit is not None else None,
        "total_credit": str(stmt.total_credit) if stmt.total_credit is not None else None,
        "currency": stmt.currency,
        "client_name": stmt.client_name,
        "client_pib": stmt.client_pib,
        "source_file": stmt.source_file,
        "status": stmt.status,
        "export_file": stmt.export_file,
        "error_message": stmt.error_message,
        "transactions": [
            {
                "id": t.id,
                "row_number": t.row_number,
                "value_date": t.value_date.isoformat() if t.value_date else None,
                "booking_date": t.booking_date.isoformat() if t.booking_date else None,
                "debit": str(t.debit) if t.debit is not None else None,
                "credit": str(t.credit) if t.credit is not None else None,
                "counterparty": t.counterparty,
                "counterparty_account": t.counterparty_account,
                "counterparty_bank": t.counterparty_bank,
                "payment_code": t.payment_code,
                "purpose": t.purpose,
                "reference_debit": t.reference_debit,
                "reference_credit": t.reference_credit,
                "reclamation_data": t.reclamation_data,
                "fee": str(t.fee) if t.fee is not None else None,
            }
            for t in stmt.transactions
        ],
    }


@app.put("/api/statement/{statement_id}")
def api_update_statement(statement_id: int, request_body: dict, db: Session = Depends(get_db)):
    stmt = db.query(Statement).filter(Statement.id == statement_id).first()
    if not stmt:
        raise HTTPException(status_code=404, detail="Statement not found")

    allowed = {
        "account_number", "iban", "statement_number", "statement_date",
        "period_start", "period_end", "opening_balance", "closing_balance",
        "total_debit", "total_credit", "currency", "client_name", "client_pib",
    }
    for key, value in request_body.items():
        if key in allowed:
            setattr(stmt, key, value)
    stmt.status = "reviewed"
    db.commit()
    return {"ok": True}


@app.put("/api/transaction/{transaction_id}")
def api_update_transaction(transaction_id: int, request_body: dict, db: Session = Depends(get_db)):
    tx = db.query(Transaction).filter(Transaction.id == transaction_id).first()
    if not tx:
        raise HTTPException(status_code=404, detail="Transaction not found")

    allowed = {
        "value_date", "booking_date", "debit", "credit",
        "counterparty", "counterparty_account", "counterparty_bank",
        "payment_code", "purpose", "reference_debit", "reference_credit",
        "reclamation_data", "fee",
    }
    for key, value in request_body.items():
        if key in allowed:
            setattr(tx, key, value)
    db.commit()
    return {"ok": True}


@app.post("/api/statement/{statement_id}/export")
def api_export(statement_id: int, db: Session = Depends(get_db)):
    stmt = db.query(Statement).filter(Statement.id == statement_id).first()
    if not stmt:
        raise HTTPException(status_code=404, detail="Statement not found")

    try:
        # Export all statements for this account into one file
        all_for_account = (
            db.query(Statement)
            .filter(
                Statement.account_number == stmt.account_number,
                Statement.status != "error",
            )
            .all()
        )
        export_path = generate_1c_file_multi(all_for_account, settings.output_dir)
        for s in all_for_account:
            s.export_file = str(export_path)
            s.status = "exported"
        db.commit()
        return {"ok": True, "file": str(export_path)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/statement/{statement_id}/download")
def api_download(statement_id: int, db: Session = Depends(get_db)):
    stmt = db.query(Statement).filter(Statement.id == statement_id).first()
    if not stmt:
        raise HTTPException(status_code=404, detail="Statement not found")
    if not stmt.export_file or not Path(stmt.export_file).exists():
        raise HTTPException(status_code=404, detail="Export file not found. Export first.")
    return FileResponse(
        stmt.export_file,
        media_type="text/plain",
        filename=Path(stmt.export_file).name,
    )


@app.delete("/api/statement/{statement_id}")
def api_delete_statement(statement_id: int, db: Session = Depends(get_db)):
    stmt = db.query(Statement).filter(Statement.id == statement_id).first()
    if not stmt:
        raise HTTPException(status_code=404, detail="Statement not found")
    db.delete(stmt)
    db.commit()
    return {"ok": True}
