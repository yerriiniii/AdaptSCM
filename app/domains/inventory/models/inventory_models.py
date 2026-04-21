import uuid
from datetime import date, datetime, timezone

from sqlalchemy import Date, DateTime, ForeignKey, Index, Integer, Numeric, String, Text, UniqueConstraint
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.shared.db.base import Base

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
    # 출고 엑셀 「판매처」 원문(탭 분류 후에도 추적). 일반 재고 업로드는 미사용.
    vendor_raw: Mapped[str | None] = mapped_column(Text, nullable=True)
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


class ProductGroup(Base):
    __tablename__ = "item"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    kr_name: Mapped[str] = mapped_column(String(255), nullable=False)
    brand: Mapped[str | None] = mapped_column(String(255), nullable=True)
    barcode: Mapped[str | None] = mapped_column(String(255), nullable=True)
    upload_updated_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    manual_updated_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True))
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False, default=_utc_now, onupdate=_utc_now
    )

    locales: Mapped[list["ProductLocale"]] = relationship(
        back_populates="product_group",
        cascade="all, delete-orphan",
    )

    __table_args__ = (
        Index("ix_item_kr_name", "kr_name"),
    )


class ProductLocale(Base):
    __tablename__ = "item_mapping"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    item_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True),
        ForeignKey("item.id", ondelete="CASCADE"),
        nullable=False,
    )
    country_code: Mapped[str] = mapped_column(String(10), nullable=False)
    name: Mapped[str | None] = mapped_column(String(255))
    sku: Mapped[str] = mapped_column(String(255), nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False, default=_utc_now, onupdate=_utc_now
    )

    product_group: Mapped["ProductGroup"] = relationship(back_populates="locales")

    __table_args__ = (
        # 국가당 SKU 여러 행 허용(예: 대만 HK06019 / HK06019-02 동일 한국 상품)
        Index("ix_item_mapping_item_country", "item_id", "country_code"),
        Index("ix_item_mapping_country_sku", "country_code", "sku", unique=True),
        Index("ix_item_mapping_country_name", "country_code", "name"),
    )


class PurchaseOrder(Base):
    """발주 헤더 (ERP PO 단위)."""

    __tablename__ = "purchase_orders"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    order_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    order_date_note: Mapped[str | None] = mapped_column(String(128), nullable=True)
    erp_po_number: Mapped[str] = mapped_column(String(128), nullable=False)
    product_type: Mapped[str] = mapped_column(String(255), nullable=False)
    sku: Mapped[str] = mapped_column(String(255), nullable=False)
    brand: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    product_name: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    manufacturer: Mapped[str] = mapped_column(String(255), nullable=False, default="")
    total_quantity: Mapped[float] = mapped_column(Numeric(18, 2), nullable=False)
    delivery_available_date: Mapped[date | None] = mapped_column(Date)
    expected_inbound_date: Mapped[date | None] = mapped_column(Date)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=_utc_now)

    inbound_lines: Mapped[list["PurchaseInboundLine"]] = relationship(
        back_populates="purchase_order",
        cascade="all, delete-orphan",
        order_by="PurchaseInboundLine.line_no",
    )

    __table_args__ = (
        UniqueConstraint(
            "erp_po_number",
            "sku",
            "order_date",
            name="uq_purchase_orders_erp_sku_order_date",
        ),
    )


class PurchaseInboundLine(Base):
    """발주별 입고 차수 (POxxxx_1, POxxxx_2 …)."""

    __tablename__ = "purchase_inbound_lines"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    purchase_order_id: Mapped[uuid.UUID] = mapped_column(
        UUID(as_uuid=True),
        ForeignKey("purchase_orders.id", ondelete="CASCADE"),
        nullable=False,
    )
    line_no: Mapped[int] = mapped_column(Integer, nullable=False)
    ref_code: Mapped[str] = mapped_column(String(160), nullable=False)
    delivery_available_date: Mapped[date | None] = mapped_column(Date)
    expected_inbound_date: Mapped[date | None] = mapped_column(Date)
    actual_inbound_date: Mapped[date | None] = mapped_column(Date)
    actual_inbound_note: Mapped[str | None] = mapped_column(String(128), nullable=True)
    line_memo: Mapped[str | None] = mapped_column(Text, nullable=True)
    quantity: Mapped[float] = mapped_column(Numeric(18, 2), nullable=False)
    inbound_status: Mapped[str] = mapped_column(String(4), nullable=False, default="O")
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=_utc_now)

    purchase_order: Mapped["PurchaseOrder"] = relationship(back_populates="inbound_lines")

    __table_args__ = (
        Index("ix_purchase_inbound_lines_order", "purchase_order_id", "line_no"),
        UniqueConstraint("purchase_order_id", "line_no", name="uq_purchase_inbound_order_line"),
    )
