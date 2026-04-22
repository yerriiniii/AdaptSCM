import uuid
from datetime import datetime, timezone

from sqlalchemy import DateTime, String
from sqlalchemy.dialects.postgresql import UUID
from sqlalchemy.orm import Mapped, mapped_column

from shared.db.base import Base


def _uuid() -> uuid.UUID:
    return uuid.uuid4()


def _utc() -> datetime:
    return datetime.now(timezone.utc)


class SignupEmailProof(Base):
    """가입 전 이메일 소유 확인용 일회성 토큰(계정 생성 전)."""

    __tablename__ = "signup_email_proofs"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid)
    email: Mapped[str] = mapped_column(String(320), nullable=False, index=True)
    token: Mapped[str] = mapped_column(String(128), nullable=False, unique=True, index=True)
    expires_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), nullable=False, default=_utc)
