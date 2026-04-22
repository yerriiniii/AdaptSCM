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


class SignupEmailSendLog(Base):
    """이메일 인증 링크 전송 횟수(24h 슬라이딩) 집계용."""

    __tablename__ = "signup_email_send_log"

    id: Mapped[uuid.UUID] = mapped_column(UUID(as_uuid=True), primary_key=True, default=_uuid)
    email: Mapped[str] = mapped_column(String(320), nullable=False, index=True)
    created_at: Mapped[datetime] = mapped_column(
        DateTime(timezone=True), nullable=False, default=_utc, index=True
    )
