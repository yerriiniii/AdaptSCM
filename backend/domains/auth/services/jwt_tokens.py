from datetime import datetime, timedelta, timezone
from typing import Any
from uuid import UUID

from jose import jwt

from shared.config import get_runtime_settings


def create_access_token(
    *,
    user_id: UUID,
    email: str,
    is_admin: bool,
    expires_at: datetime | None = None,
) -> tuple[str, int]:
    """Returns (token, expires_in_seconds)."""
    s = get_runtime_settings()
    now = datetime.now(timezone.utc)
    exp = expires_at or now + timedelta(minutes=s.jwt_expire_minutes)
    payload: dict[str, Any] = {
        "sub": str(user_id),
        "email": email,
        "adm": is_admin,
        "iat": int(now.timestamp()),
        "exp": int(exp.timestamp()),
    }
    token = jwt.encode(payload, s.jwt_secret, algorithm=s.jwt_algorithm)
    expires_in = max(0, int(exp.timestamp() - now.timestamp()))
    return token, expires_in


def decode_access_token(token: str) -> dict[str, Any]:
    s = get_runtime_settings()
    return jwt.decode(token, s.jwt_secret, algorithms=[s.jwt_algorithm])
