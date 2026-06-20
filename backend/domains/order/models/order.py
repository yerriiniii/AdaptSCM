import uuid
from datetime import date, datetime, timezone

from sqlalchemy import Date, DateTime, ForeignKey, Index, Integer, Numeric, String, Text, UniqueConstraint
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from shared.db.base import Base


def _uuid_value() -> uuid.UUID:
    return uuid.uuid4()


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


class PurchaseOrder(Base):
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
