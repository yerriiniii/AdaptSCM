from fastapi import APIRouter, Depends, HTTPException, Query, status
from pydantic import EmailStr
from sqlalchemy.orm import Session

from domains.auth.deps import require_active_user
from domains.auth.dto.auth_dto import (
    AccountStatusResponse,
    ChangePasswordRequest,
    LoginRequest,
    MeResponse,
    MessageResponse,
    RegisterRequest,
    RegisterResponse,
    SignupEmailSendRequest,
    SignupEmailSendResponse,
    SignupTokenCheckResponse,
    TokenResponse,
)
from domains.auth.models.user_models import UserAccount, UserAccountStatus
from domains.auth.services.jwt_tokens import create_access_token
from domains.auth.services.password_hashing import verify_password
from domains.auth.services.signup_email_proof_service import (
    check_proof_token,
    create_or_refresh_proof_and_send_email,
    register_user_with_proof,
)
from domains.auth.services.signup_email_send_rate_service import count_sends_24h, is_rate_limited, log_send
from domains.auth.services.user_account_service import change_user_password, get_user_by_email
from shared.db import get_db_session

router = APIRouter(prefix="/api/auth", tags=["auth"])


@router.post("/signup-email/send", response_model=SignupEmailSendResponse)
def signup_email_send(
    body: SignupEmailSendRequest,
    db: Session = Depends(get_db_session),
) -> SignupEmailSendResponse:
    """가입 1단계: 이메일로 인증 링크 발송(계정은 아직 만들지 않음)."""
    email = str(body.email).strip().lower()
    if get_user_by_email(db, email) is not None:
        raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="이미 등록된 이메일입니다.")
    if is_rate_limited(db, email):
        raise HTTPException(
            status_code=status.HTTP_429_TOO_MANY_REQUESTS,
            detail="인증 링크는 24시간에 최대 5번까지 보낼 수 있습니다. 24시간이 지난 뒤 다시 시도해 주세요.",
        )
    _proof, signup_url, mail_sent = create_or_refresh_proof_and_send_email(db, email)
    log_send(db, email)
    n = count_sends_24h(db, email)
    ok_msg = (
        "인증 링크를 메일로 보냈습니다.\n"
        "메일에 첨부된 링크를 눌러 비밀번호를 입력하고 회원가입을 요청하세요."
    )
    if mail_sent:
        return SignupEmailSendResponse(
            message=ok_msg,
            mail_sent=True,
            signup_url=None,
            send_count_24h=n,
            expires_at=_proof.expires_at,
        )
    return SignupEmailSendResponse(
        message="인증 링크를 발급했습니다. 아래 링크로 이동해 비밀번호를 입력하고 회원가입을 요청하세요.",
        mail_sent=False,
        signup_url=signup_url,
        send_count_24h=n,
        expires_at=_proof.expires_at,
    )


@router.get("/signup-email/check", response_model=SignupTokenCheckResponse)
def signup_email_check(
    token: str = Query(..., min_length=20, description="이메일 링크의 signup_token"),
    db: Session = Depends(get_db_session),
) -> SignupTokenCheckResponse:
    """링크에 포함된 토큰으로 이메일·유효성 확인(가입 폼을 띄울 때 호출)."""
    p = check_proof_token(db, token)
    if p is None:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="유효하지 않거나 만료된 링크입니다. 인증 메일을 다시 요청해 주세요.",
        )
    return SignupTokenCheckResponse(email=p.email, valid=True)


@router.post(
    "/register",
    response_model=RegisterResponse,
    status_code=status.HTTP_201_CREATED,
)
def register(body: RegisterRequest, db: Session = Depends(get_db_session)) -> RegisterResponse:
    """가입 2단계: 이메일 인증 토큰 + 비밀번호(계정 생성, pending, 이메일 인증됨)."""
    email = str(body.email).strip().lower()
    try:
        register_user_with_proof(
            db,
            email=email,
            password=body.password,
            signup_token=body.signup_token,
        )
    except ValueError as exc:
        code = exc.args[0] if exc.args else ""
        if code == "INVALID_SIGNUP_TOKEN":
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=(
                    "유효하지 않거나 만료된 가입 토큰입니다.\n"
                    "이메일의 링크로 다시 들어오거나 인증 메일을 다시 보내 주세요."
                ),
            ) from exc
        if code == "EMAIL_MISMATCH":
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="이메일이 인증 링크의 주소와 일치하지 않습니다.",
            ) from exc
        if code == "EMAIL_TAKEN":
            raise HTTPException(status_code=status.HTTP_409_CONFLICT, detail="이미 등록된 이메일입니다.") from exc
        raise
    return RegisterResponse(
        message="가입 요청이 접수되었습니다. 관리자 승인 후 로그인할 수 있습니다.",
        status=UserAccountStatus.PENDING.value,
    )


@router.get("/account-status", response_model=AccountStatusResponse)
def account_status(
    email: EmailStr = Query(..., description="조회할 계정 이메일"),
    db: Session = Depends(get_db_session),
) -> AccountStatusResponse:
    """이메일로 계정 승인 상태만 조회(로그인 전 안내 UI용)."""
    em = str(email).strip().lower()
    u = get_user_by_email(db, em)
    if u is None:
        return AccountStatusResponse(status="none")
    if u.status == UserAccountStatus.REJECTED.value:
        return AccountStatusResponse(status="rejected")
    if u.status == UserAccountStatus.ACTIVE.value:
        return AccountStatusResponse(status="active")
    if u.status == UserAccountStatus.PENDING.value:
        return AccountStatusResponse(status="pending")
    return AccountStatusResponse(status="none")


@router.post("/login", response_model=TokenResponse)
def login(body: LoginRequest, db: Session = Depends(get_db_session)) -> TokenResponse:
    email = str(body.email).strip().lower()
    u = get_user_by_email(db, email)
    if u is None or not verify_password(body.password, u.password_hash):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="이메일 또는 비밀번호가 올바르지 않습니다.")
    if not u.email_verified:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail={"message": "이메일 인증이 완료되지 않은 계정입니다. 관리자에게 문의하세요.", "code": "email_not_verified"},
        )
    if u.status == UserAccountStatus.PENDING.value:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail={"message": "관리자 승인 대기 중입니다.", "code": "pending_approval"},
        )
    if u.status == UserAccountStatus.REJECTED.value:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail={"message": "가입이 거절된 계정입니다. 관리자에게 문의하세요.", "code": "rejected"},
        )
    if u.status != UserAccountStatus.ACTIVE.value:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="이 계정은 로그인할 수 없는 상태입니다.",
        )
    token, expires_in = create_access_token(user_id=u.id, email=u.email, is_admin=u.is_admin)
    return TokenResponse(access_token=token, expires_in=expires_in)


@router.get("/me", response_model=MeResponse)
def me(user: UserAccount = Depends(require_active_user)) -> MeResponse:
    return MeResponse(
        id=user.id,
        email=user.email,
        status=user.status,
        is_admin=user.is_admin,
        email_verified=user.email_verified,
        created_at=user.created_at,
        approved_at=user.approved_at,
    )


@router.post("/change-password", response_model=MessageResponse)
def change_password(
    body: ChangePasswordRequest,
    db: Session = Depends(get_db_session),
    user: UserAccount = Depends(require_active_user),
) -> MessageResponse:
    try:
        change_user_password(
            db,
            user,
            current_password=body.current_password,
            new_password=body.new_password,
        )
    except ValueError as exc:
        code = exc.args[0] if exc.args else ""
        if code == "WRONG_CURRENT_PASSWORD":
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="현재 비밀번호가 올바르지 않습니다.",
            ) from exc
        raise
    return MessageResponse(message="비밀번호가 변경되었습니다. 다음 로그인부터 새 비밀번호를 사용하세요.")
