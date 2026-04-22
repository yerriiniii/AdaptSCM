import enum
import uuid
from datetime import datetime, timezone

from sqlalchemy import Boolean, DateTime, ForeignKey, String, UniqueConstraint
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column, relationship

from shared.db.base import Base


def _uuid_value() -> uuid.UUID:
    return uuid.uuid4()


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)


class UserAccountStatus(str, enum.Enum):
    PENDING = "pending"
    ACTIVE = "active"
    REJECTED = "rejected"


class UserAccount(Base):
    """승인제 가입: pending → 관리자 승인 시 active."""

    __tablename__ = "user_accounts"
    __table_args__ = (UniqueConstraint("email", name="uq_user_accounts_email"),)

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid_value)
    email: Mapped[str] = mapped_column(String(320), nullable=False, index=True)
    password_hash: Mapped[str] = mapped_column(String(255), nullable=False)
    status: Mapped[str] = mapped_column(String(20), nullable=False, default=UserAccountStatus.PENDING.value)
    is_admin: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=_utc_now)
    approved_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    approved_by_id: Mapped[uuid.UUID | None] = mapped_column(
        UUID(as_uuid=True), ForeignKey("user_accounts.id", ondelete="SET NULL"), nullable=True
    )
    approver: Mapped["UserAccount | None"] = relationship(
        "UserAccount",
        remote_side="UserAccount.id",
        foreign_keys=[approved_by_id],
    )
    # 이메일 인증(가입 링크). 인증 전에는 로그인·관리자 승인 불가.
    email_verified: Mapped[bool] = mapped_column(Boolean, nullable=False, default=False)
    email_verification_token: Mapped[str | None] = mapped_column(String(128), nullable=True, index=True)
    email_verification_expires_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
