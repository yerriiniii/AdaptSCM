import uuid
from datetime import date, datetime, timezone

from sqlalchemy import Date, DateTime, ForeignKey, Index, String
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from shared.db.base import Base


def _uuid_value() -> uuid.UUID:
    return uuid.uuid4()


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


class ProductGroup(Base):
    __tablename__ = "item"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    kr_name: Mapped[str] = mapped_column(String(255), nullable=False)
    brand: Mapped[str | None] = mapped_column(String(255), nullable=True)
    barcode: Mapped[str | None] = mapped_column(String(255), nullable=True)
    segment: Mapped[str | None] = mapped_column(String(255), nullable=True)
    mkt_priority: Mapped[str | None] = mapped_column(String(255), nullable=True)
    version: Mapped[str | None] = mapped_column(String(64), nullable=True)
    stock_category: Mapped[str | None] = mapped_column(String(64), nullable=True)
    fcst_grade: Mapped[str | None] = mapped_column(String(32), nullable=True)
    stock_grade: Mapped[str | None] = mapped_column(String(32), nullable=True)
    release_date: Mapped[date | None] = mapped_column(Date, nullable=True)
    release_month: Mapped[str | None] = mapped_column(String(64), nullable=True)
    code_registered_at: Mapped[str | None] = mapped_column(String(64), nullable=True)
    us_grade: Mapped[str | None] = mapped_column(String(32), nullable=True)
    tw_grade: Mapped[str | None] = mapped_column(String(32), nullable=True)
    hk_grade: Mapped[str | None] = mapped_column(String(32), nullable=True)
    jp_grade: Mapped[str | None] = mapped_column(String(32), nullable=True)
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
        Index("ix_item_mapping_item_country", "item_id", "country_code"),
        Index("ix_item_mapping_country_sku", "country_code", "sku", unique=True),
        Index("ix_item_mapping_country_name", "country_code", "name"),
    )


__all__ = ["ProductGroup", "ProductLocale"]
