from uuid import UUID

from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from jose import JWTError
from sqlalchemy.orm import Session

from domains.auth.models.user_models import UserAccount, UserAccountStatus
from domains.auth.services.jwt_tokens import decode_access_token
from domains.auth.services.user_account import get_user_by_id
from shared.db import get_db_session

_bearer = HTTPBearer(auto_error=True)


def _get_token_payload(credentials: HTTPAuthorizationCredentials = Depends(_bearer)) -> dict:
    if credentials.scheme.lower() != "bearer":
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="잘못된 인증 방식입니다.")
    try:
        return decode_access_token(credentials.credentials)
    except JWTError as exc:  # noqa: BLE001
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="유효하지 않은 토큰입니다.") from exc


def get_current_user(
    payload: dict = Depends(_get_token_payload),
    db: Session = Depends(get_db_session),
) -> UserAccount:
    sub = payload.get("sub")
    if not sub:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="유효하지 않은 토큰입니다.")
    try:
        user_id = UUID(str(sub))
    except (TypeError, ValueError) as exc:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="유효하지 않은 토큰입니다.") from exc
    user = get_user_by_id(db, user_id)
    if user is None:
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="사용자를 찾을 수 없습니다.")
    return user


def require_active_user(user: UserAccount = Depends(get_current_user)) -> UserAccount:
    if user.status != UserAccountStatus.ACTIVE.value:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="이 계정은 서비스를 사용할 수 없는 상태입니다.",
        )
    return user


__all__ = ["get_current_user", "require_active_user"]
