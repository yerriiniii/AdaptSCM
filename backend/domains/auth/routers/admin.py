from uuid import UUID

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.orm import Session

from domains.auth.dependencies import require_admin
from domains.auth.models.user_models import UserAccount, UserAccountStatus
from domains.auth.schemas.auth import MessageResponse, PendingUserItem, PendingUsersResponse
from domains.auth.services.user_account import approve_user, get_user_by_id, list_pending_users, reject_user
from shared.db import get_db_session

router = APIRouter(prefix="/api/auth/admin", tags=["auth-admin"])


@router.get("/pending-users", response_model=PendingUsersResponse)
def pending_users(
    db: Session = Depends(get_db_session),
    _admin: UserAccount = Depends(require_admin),
) -> PendingUsersResponse:
    items = [
        PendingUserItem(
            id=u.id,
            email=u.email,
            status=u.status,
            email_verified=u.email_verified,
            created_at=u.created_at,
        )
        for u in list_pending_users(db)
    ]
    return PendingUsersResponse(items=items)


@router.post("/users/{user_id}/approve", response_model=MessageResponse)
def approve(
    user_id: UUID,
    db: Session = Depends(get_db_session),
    admin: UserAccount = Depends(require_admin),
) -> MessageResponse:
    cand = get_user_by_id(db, user_id)
    if cand is None or cand.status != UserAccountStatus.PENDING.value:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="승인할 대기 중인 사용자를 찾을 수 없습니다.")
    if not cand.email_verified:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="이메일 인증이 완료된 계정만 승인할 수 있습니다.",
        )
    u = approve_user(db, user_id, admin)
    if u is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="승인할 대기 중인 사용자를 찾을 수 없습니다.")
    return MessageResponse(message="승인되었습니다.")


@router.post("/users/{user_id}/reject", response_model=MessageResponse)
def reject(
    user_id: UUID,
    db: Session = Depends(get_db_session),
    admin: UserAccount = Depends(require_admin),
) -> MessageResponse:
    u = reject_user(db, user_id, admin)
    if u is None:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="거절할 대기 중인 사용자를 찾을 수 없습니다.")
    return MessageResponse(message="거절되었습니다.")

__all__ = ["router"]
