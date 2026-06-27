import logging
import secrets

from sqlalchemy import select
from sqlalchemy.orm import Session

from domains.auth.models.user_models import UserAccount, UserAccountStatus
from domains.auth.services.password_hashing import hash_password

_log = logging.getLogger(__name__)

GATE_ACCESS_USER_EMAIL = "__access_gate__@system.local"


def get_user_by_email(db: Session, email: str) -> UserAccount | None:
    q = select(UserAccount).where(UserAccount.email == email)
    return db.execute(q).scalar_one_or_none()


def get_user_by_id(db: Session, user_id) -> UserAccount | None:
    return db.get(UserAccount, user_id)


def ensure_gate_access_user(db: Session) -> UserAccount:
    email = GATE_ACCESS_USER_EMAIL
    existing = get_user_by_email(db, email)
    if existing is not None:
        if existing.status != UserAccountStatus.ACTIVE.value:
            existing.status = UserAccountStatus.ACTIVE.value
            existing.is_admin = True
            db.add(existing)
            db.commit()
            db.refresh(existing)
        return existing
    u = UserAccount(
        email=email,
        password_hash=hash_password(secrets.token_urlsafe(48)),
        status=UserAccountStatus.ACTIVE.value,
        is_admin=True,
    )
    db.add(u)
    db.commit()
    db.refresh(u)
    _log.info("관리자 인증번호 접근용 시스템 계정을 생성했습니다.")
    return u
