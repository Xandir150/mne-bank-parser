from datetime import date, datetime
from decimal import Decimal
from typing import Optional

from sqlalchemy import ForeignKey, String, Text, Numeric, Date, DateTime, Integer
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.database import Base


class Statement(Base):
    __tablename__ = "statements"

    id: Mapped[int] = mapped_column(primary_key=True)
    bank_code: Mapped[str] = mapped_column(String(10))
    bank_name: Mapped[str] = mapped_column(String(100))
    account_number: Mapped[str] = mapped_column(String(50))
    iban: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    statement_number: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    statement_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    period_start: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    period_end: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    opening_balance: Mapped[Optional[Decimal]] = mapped_column(Numeric(15, 2), nullable=True)
    closing_balance: Mapped[Optional[Decimal]] = mapped_column(Numeric(15, 2), nullable=True)
    total_debit: Mapped[Optional[Decimal]] = mapped_column(Numeric(15, 2), nullable=True)
    total_credit: Mapped[Optional[Decimal]] = mapped_column(Numeric(15, 2), nullable=True)
    currency: Mapped[str] = mapped_column(String(10), default="EUR")
    client_name: Mapped[Optional[str]] = mapped_column(String(200), nullable=True)
    client_pib: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    source_file: Mapped[str] = mapped_column(String(500))
    status: Mapped[str] = mapped_column(String(20), default="new")
    export_file: Mapped[Optional[str]] = mapped_column(String(500), nullable=True)
    error_message: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, default=datetime.utcnow, onupdate=datetime.utcnow
    )

    transactions: Mapped[list["Transaction"]] = relationship(
        back_populates="statement", cascade="all, delete-orphan", order_by="Transaction.row_number"
    )


class Transaction(Base):
    __tablename__ = "transactions"

    id: Mapped[int] = mapped_column(primary_key=True)
    statement_id: Mapped[int] = mapped_column(ForeignKey("statements.id"))
    row_number: Mapped[int] = mapped_column(Integer, default=0)
    value_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    booking_date: Mapped[Optional[date]] = mapped_column(Date, nullable=True)
    debit: Mapped[Optional[Decimal]] = mapped_column(Numeric(15, 2), nullable=True)
    credit: Mapped[Optional[Decimal]] = mapped_column(Numeric(15, 2), nullable=True)
    counterparty: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    counterparty_account: Mapped[Optional[str]] = mapped_column(String(50), nullable=True)
    counterparty_bank: Mapped[Optional[str]] = mapped_column(String(200), nullable=True)
    payment_code: Mapped[Optional[str]] = mapped_column(String(20), nullable=True)
    purpose: Mapped[Optional[str]] = mapped_column(Text, nullable=True)
    reference_debit: Mapped[Optional[str]] = mapped_column(String(100), nullable=True)
    reference_credit: Mapped[Optional[str]] = mapped_column(String(100), nullable=True)
    reclamation_data: Mapped[Optional[str]] = mapped_column(String(200), nullable=True)
    fee: Mapped[Optional[Decimal]] = mapped_column(Numeric(15, 2), nullable=True)

    statement: Mapped["Statement"] = relationship(back_populates="transactions")
