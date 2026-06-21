import logging
import secrets
from datetime import datetime, timezone
from uuid import UUID

from sqlalchemy import func, select
from sqlalchemy.orm import Session

from domains.auth.models.user_models import UserAccount, UserAccountStatus
from domains.auth.services.password_hashing import hash_password, verify_password
from shared.config import get_runtime_settings

_log = logging.getLogger(__name__)

GATE_ACCESS_USER_EMAIL = "__access_gate__@system.local"


def get_user_by_email(db: Session, email: str) -> UserAccount | None:
    q = select(UserAccount).where(UserAccount.email == email)
    return db.execute(q).scalar_one_or_none()


def get_user_by_id(db: Session, user_id: UUID) -> UserAccount | None:
    return db.get(UserAccount, user_id)


def change_user_password(
    db: Session,
    user: UserAccount,
    *,
    current_password: str,
    new_password: str,
) -> None:
    if not verify_password(current_password, user.password_hash):
        raise ValueError("WRONG_CURRENT_PASSWORD")
    user.password_hash = hash_password(new_password)
    db.add(user)
    db.commit()
    db.refresh(user)


def ensure_gate_access_user(db: Session) -> UserAccount:
    email = GATE_ACCESS_USER_EMAIL
    existing = get_user_by_email(db, email)
    if existing is not None:
        if existing.status != UserAccountStatus.ACTIVE.value:
            existing.status = UserAccountStatus.ACTIVE.value
            existing.is_admin = True
            existing.email_verified = True
            db.add(existing)
            db.commit()
            db.refresh(existing)
        return existing
    u = UserAccount(
        email=email,
        password_hash=hash_password(secrets.token_urlsafe(48)),
        status=UserAccountStatus.ACTIVE.value,
        is_admin=True,
        email_verified=True,
    )
    db.add(u)
    db.commit()
    db.refresh(u)
    _log.info("관리자 인증번호 접근용 시스템 계정을 생성했습니다.")
    return u


def ensure_initial_admin_user(db: Session) -> None:
    s = get_runtime_settings()
    if not s.initial_admin_email or not s.initial_admin_password:
        return
    email = s.initial_admin_email.strip().lower()
    if not email:
        return
    existing = get_user_by_email(db, email)
    if existing is not None:
        return
    u = UserAccount(
        email=email,
        password_hash=hash_password(s.initial_admin_password),
        status=UserAccountStatus.ACTIVE.value,
        is_admin=True,
        email_verified=True,
    )
    db.add(u)
    db.commit()
    _log.info("초기 관리자 계정을 생성했습니다(이메일 1회).")


def list_pending_users(db: Session) -> list[UserAccount]:
    q = (
        select(UserAccount)
        .where(UserAccount.status == UserAccountStatus.PENDING.value)
        .order_by(UserAccount.created_at.asc())
    )
    return list(db.execute(q).scalars().all())


def approve_user(db: Session, user_id: UUID, approver: UserAccount) -> UserAccount | None:
    u = get_user_by_id(db, user_id)
    if u is None or u.status != UserAccountStatus.PENDING.value:
        return None

    u.status = UserAccountStatus.ACTIVE.value
    u.approved_at = datetime.now(timezone.utc)
    u.approved_by_id = approver.id
    db.add(u)
    db.commit()
    db.refresh(u)
    return u


def reject_user(db: Session, user_id: UUID, _actor: UserAccount) -> UserAccount | None:
    u = get_user_by_id(db, user_id)
    if u is None or u.status != UserAccountStatus.PENDING.value:
        return None
    u.status = UserAccountStatus.REJECTED.value
    u.approved_at = None
    u.approved_by_id = None
    db.add(u)
    db.commit()
    db.refresh(u)
    return u


def count_other_active_admins(db: Session, exclude_id: UUID) -> int:
    q = (
        select(func.count())
        .select_from(UserAccount)
        .where(
            UserAccount.is_admin.is_(True),
            UserAccount.status == UserAccountStatus.ACTIVE.value,
            UserAccount.id != exclude_id,
        )
    )
    return int(db.execute(q).scalar_one() or 0)


def delete_own_account(db: Session, user: UserAccount) -> None:
    if user.is_admin and count_other_active_admins(db, user.id) < 1:
        raise ValueError("SOLE_ACTIVE_ADMIN")
    db.delete(user)
    db.commit()
