from datetime import datetime, timedelta, timezone

from fastapi import APIRouter, Depends
from sqlalchemy.orm import Session

from domains.auth.dependencies import get_token_payload, require_active_user
from domains.auth.models.user_models import UserAccount
from domains.auth.schemas.auth import AdminAccessCodeRequest, MeResponse, TokenResponse
from domains.auth.services.admin_access import verify_admin_access_code
from domains.auth.services.jwt_tokens import create_access_token
from domains.auth.services.user_account import GATE_ACCESS_USER_EMAIL, ensure_gate_access_user
from shared.db import get_db_session

router = APIRouter(prefix="/api/auth", tags=["auth"])

SESSION_EXTENSION_HOURS = 3


@router.post("/access-code", response_model=TokenResponse)
def access_code_login(
    body: AdminAccessCodeRequest,
    db: Session = Depends(get_db_session),
) -> TokenResponse:
    verify_admin_access_code(body.code)
    user = ensure_gate_access_user(db)
    token, expires_in = create_access_token(user_id=user.id, email=user.email, is_admin=True)
    return TokenResponse(access_token=token, expires_in=expires_in)


@router.post("/extend", response_model=TokenResponse)
def extend_session(
    user: UserAccount = Depends(require_active_user),
    payload: dict = Depends(get_token_payload),
) -> TokenResponse:
    now = datetime.now(timezone.utc)
    current_exp_ts = payload.get("exp")
    current_exp = (
        datetime.fromtimestamp(current_exp_ts, timezone.utc)
        if isinstance(current_exp_ts, (int, float))
        else now
    )
    expires_at = max(current_exp, now) + timedelta(hours=SESSION_EXTENSION_HOURS)
    token, expires_in = create_access_token(
        user_id=user.id,
        email=user.email,
        is_admin=user.is_admin,
        expires_at=expires_at,
    )
    return TokenResponse(access_token=token, expires_in=expires_in)


@router.get("/me", response_model=MeResponse)
def me(user: UserAccount = Depends(require_active_user)) -> MeResponse:
    display_email = "관리자" if user.email == GATE_ACCESS_USER_EMAIL else user.email
    return MeResponse(
        id=user.id,
        email=display_email,
        status=user.status,
        is_admin=user.is_admin,
        created_at=user.created_at,
    )


__all__ = ["router"]
