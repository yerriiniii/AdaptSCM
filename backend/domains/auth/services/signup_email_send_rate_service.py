from datetime import datetime, timedelta, timezone

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from domains.auth.models.signup_email_send_log_model import SignupEmailSendLog

_MAX_SENDS_24H = 5
_WINDOW = timedelta(hours=24)


def _now() -> datetime:
    return datetime.now(timezone.utc)


def count_sends_24h(db: Session, email: str) -> int:
    email = (email or "").strip().lower()
    cutoff = _now() - _WINDOW
    q = select(func.count()).select_from(SignupEmailSendLog).where(
        SignupEmailSendLog.email == email,
        SignupEmailSendLog.created_at >= cutoff,
    )
    return int(db.execute(q).scalar() or 0)


def log_send(db: Session, email: str) -> None:
    email = (email or "").strip().lower()
    db.add(SignupEmailSendLog(email=email))
    db.commit()


def is_rate_limited(db: Session, email: str) -> bool:
    return count_sends_24h(db, email) >= _MAX_SENDS_24H
