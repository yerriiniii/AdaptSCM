import uuid
from datetime import datetime, timezone

from sqlalchemy import DateTime, Index, Integer, JSON, String, UniqueConstraint
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column

from shared.db.base import Base


def _uuid_value() -> uuid.UUID:
    return uuid.uuid4()


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


class ShipmentMatrixRow(Base):
    __tablename__ = "shipment_matrix_rows"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    channel_sheet: Mapped[str] = mapped_column(String(128), nullable=False)
    data_year: Mapped[int] = mapped_column(Integer, nullable=False, default=2026)
    sku: Mapped[str] = mapped_column(String(255), nullable=False)
    mkt_priority: Mapped[str | None] = mapped_column(String(255), nullable=True)
    category: Mapped[str | None] = mapped_column(String(255), nullable=True)
    monthly_totals: Mapped[dict | None] = mapped_column(JSON, nullable=True)
    daily_totals: Mapped[dict | None] = mapped_column(JSON, nullable=True)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False, default=_utc_now, onupdate=_utc_now
    )

    __table_args__ = (
        UniqueConstraint("channel_sheet", "data_year", "sku", name="uq_shipment_matrix_channel_year_sku"),
        Index("ix_shipment_matrix_channel_year", "channel_sheet", "data_year"),
    )
