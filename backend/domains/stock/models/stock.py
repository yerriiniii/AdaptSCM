import uuid
from datetime import date, datetime, timezone

from sqlalchemy import Date, DateTime, ForeignKey, Index, Integer, Numeric, String, Text
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from shared.db.base import Base


def _uuid_value() -> uuid.UUID:
    return uuid.uuid4()


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


class UploadedFile(Base):
    __tablename__ = "uploaded_files"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    original_name: Mapped[str] = mapped_column(String(255), nullable=False)
    stored_name: Mapped[str] = mapped_column(String(255), nullable=False)
    s3_bucket: Mapped[str | None] = mapped_column(String(255))
    s3_key: Mapped[str | None] = mapped_column(Text)
    country_type: Mapped[str] = mapped_column(String(20), nullable=False)
    country_code: Mapped[str] = mapped_column(String(10), nullable=False)
    base_date: Mapped[date | None] = mapped_column(Date)
    uploaded_by: Mapped[str | None] = mapped_column(String(100))
    size: Mapped[int | None] = mapped_column(Integer)
    status: Mapped[str] = mapped_column(String(30), nullable=False, default="uploaded")
    error_message: Mapped[str | None] = mapped_column(Text)
    uploaded_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=_utc_now)
    file_domain: Mapped[str] = mapped_column(String(20), nullable=False, default="inventory")
    shipment_sheet_key: Mapped[str | None] = mapped_column(String(64), nullable=True)

    inventory_rows: Mapped[list["InventoryRow"]] = relationship(
        back_populates="uploaded_file",
        cascade="all, delete-orphan",
    )

    __table_args__ = (
        Index("ix_uploaded_files_country_scope", "country_type", "country_code", "base_date"),
        Index("ix_uploaded_files_file_domain", "file_domain"),
    )


class InventoryRow(Base):
    __tablename__ = "inventory_rows"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    uploaded_file_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True),
        ForeignKey("uploaded_files.id", ondelete="CASCADE"),
        nullable=False,
    )
    sku: Mapped[str] = mapped_column(String(255), nullable=False)
    description: Mapped[str | None] = mapped_column(String(255))
    supplier: Mapped[str | None] = mapped_column(String(255))
    level: Mapped[str | None] = mapped_column(String(20))
    warehouse: Mapped[str | None] = mapped_column(String(255))
    vendor_raw: Mapped[str | None] = mapped_column(Text, nullable=True)
    shipment_order_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=False), nullable=True)
    quantity: Mapped[float] = mapped_column(Numeric(18, 2), nullable=False, default=0)
    uploaded_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=_utc_now)
    row_snapshot_date: Mapped[date | None] = mapped_column(Date)
    row_country_code: Mapped[str | None] = mapped_column(String(10))

    uploaded_file: Mapped["UploadedFile"] = relationship(back_populates="inventory_rows")

    __table_args__ = (
        Index("ix_inventory_rows_uploaded_file_id", "uploaded_file_id"),
        Index("ix_inventory_rows_upload_snapshot", "uploaded_file_id", "row_snapshot_date"),
        Index("ix_inventory_rows_sku", "sku"),
    )


class InventoryAggregate(Base):
    __tablename__ = "inventory_aggregates"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    country_code: Mapped[str] = mapped_column(String(10), nullable=False)
    base_date: Mapped[date | None] = mapped_column(Date)
    sku: Mapped[str] = mapped_column(String(255), nullable=False)
    description: Mapped[str | None] = mapped_column(String(255))
    supplier: Mapped[str | None] = mapped_column(String(255))
    level: Mapped[str | None] = mapped_column(String(20))
    warehouse: Mapped[str | None] = mapped_column(String(255))
    total_quantity: Mapped[float] = mapped_column(Numeric(18, 2), nullable=False, default=0)
    aggregation_version: Mapped[int] = mapped_column(Integer, nullable=False)
    uploaded_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=_utc_now)
    source_domain: Mapped[str] = mapped_column(String(20), nullable=False, default="INVENTORY")

    __table_args__ = (
        Index("ix_inventory_aggregates_country_date", "country_code", "base_date"),
        Index("ix_inventory_aggregates_sku", "sku"),
        Index("ix_inventory_aggregates_version", "aggregation_version"),
    )
