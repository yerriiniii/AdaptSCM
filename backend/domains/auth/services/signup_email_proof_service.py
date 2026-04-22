import logging
import secrets
from datetime import datetime, timedelta, timezone

from sqlalchemy import delete, select
from sqlalchemy.orm import Session

from domains.auth.models.signup_email_proof_model import SignupEmailProof
from domains.auth.models.user_models import UserAccount, UserAccountStatus
from domains.auth.services.auth_email_service import send_signup_email_proof_link
from domains.auth.services.password_hashing import hash_password
from shared.config import get_runtime_settings

_log = logging.getLogger(__name__)


def _now() -> datetime:
    return datetime.now(timezone.utc)


def get_proof_by_token(db: Session, token: str) -> SignupEmailProof | None:
    t = (token or "").strip()
    if not t:
        return None
    q = select(SignupEmailProof).where(SignupEmailProof.token == t)
    return db.execute(q).scalar_one_or_none()


def create_or_refresh_proof_and_send_email(
    db: Session, email: str
) -> tuple[SignupEmailProof, str, bool]:
    """
    기존 동일 이메일 proof 행은 삭제하고 새 토큰을 발급, 메일 발송.
    (email, email_sent)에서 email은 항상 반환(발송 실패해도 토큰은 DB에 있음)
    """
    s = get_runtime_settings()
    email = email.strip().lower()
    db.execute(delete(SignupEmailProof).where(SignupEmailProof.email == email))
    db.commit()

    raw = secrets.token_urlsafe(32)
    exp = _now() + timedelta(seconds=s.email_verification_ttl_sec)
    row = SignupEmailProof(email=email, token=raw, expires_at=exp)
    db.add(row)
    db.commit()
    db.refresh(row)

    url = f"{s.public_app_url}/?signup_token={raw}"
    mail_sent = send_signup_email_proof_link(to_email=email, signup_url=url)
    return row, url, mail_sent


def check_proof_token(db: Session, token: str) -> SignupEmailProof | None:
    p = get_proof_by_token(db, token)
    if p is None:
        return None
    if p.expires_at < _now():
        return None
    return p


def register_user_with_proof(
    db: Session,
    *,
    email: str,
    password: str,
    signup_token: str,
) -> UserAccount:
    email = email.strip().lower()
    p = check_proof_token(db, signup_token)
    if p is None:
        raise ValueError("INVALID_SIGNUP_TOKEN")
    if p.email != email:
        raise ValueError("EMAIL_MISMATCH")
    u = (
        db.execute(select(UserAccount).where(UserAccount.email == email))
        .scalar_one_or_none()
    )
    if u is not None:
        raise ValueError("EMAIL_TAKEN")

    proof_id = p.id
    user = UserAccount(
        email=email,
        password_hash=hash_password(password),
        status=UserAccountStatus.PENDING.value,
        is_admin=False,
        email_verified=True,
    )
    db.add(user)
    db.execute(delete(SignupEmailProof).where(SignupEmailProof.id == proof_id))
    db.commit()
    db.refresh(user)
    return user
