from datetime import datetime
from typing import Literal
from uuid import UUID

from pydantic import BaseModel, EmailStr, Field, field_validator, model_validator

from domains.auth.services.password_policy import validate_password_strength


class SignupEmailSendRequest(BaseModel):
    email: EmailStr


class SignupEmailSendResponse(BaseModel):
    message: str
    mail_sent: bool
    """True면 SMTP로 발송이 끝남. False면 `signup_url`로 직접 열기(로컬·SMTP오류)."""
    signup_url: str | None = None
    send_count_24h: int = 0
    """최근 24시간 동안 이 이메일로 보낸 누적 횟수(이번 요청 포함)."""
    expires_at: datetime
    """이번에 발급한 인증 토큰(링크) 만료 시각(UTC, 서버 기준)."""


class SignupTokenCheckResponse(BaseModel):
    email: str
    valid: bool = True


class RegisterRequest(BaseModel):
    email: EmailStr
    password: str = Field(min_length=8, max_length=16)
    password_confirm: str = Field(min_length=8, max_length=16)
    signup_token: str = Field(min_length=20, max_length=200, description="이메일 인증 링크에 포함된 토큰")

    @field_validator("password")
    @classmethod
    def _password_policy(cls, v: str) -> str:
        err = validate_password_strength(v)
        if err:
            raise ValueError(err)
        return v

    @model_validator(mode="after")
    def _passwords_match(self) -> "RegisterRequest":
        if self.password != self.password_confirm:
            raise ValueError("비밀번호가 일치하지 않습니다.")
        return self


class LoginRequest(BaseModel):
    email: EmailStr
    password: str = Field(min_length=1, max_length=128)


class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    expires_in: int  # seconds


class MeResponse(BaseModel):
    id: UUID
    email: str
    status: str
    is_admin: bool
    email_verified: bool
    created_at: datetime
    approved_at: datetime | None


class ChangePasswordRequest(BaseModel):
    current_password: str = Field(min_length=1, max_length=128)
    new_password: str = Field(min_length=8, max_length=16)
    new_password_confirm: str = Field(min_length=8, max_length=16)

    @field_validator("new_password")
    @classmethod
    def _new_password_policy(cls, v: str) -> str:
        err = validate_password_strength(v)
        if err:
            raise ValueError(err)
        return v

    @model_validator(mode="after")
    def _new_passwords_match(self) -> "ChangePasswordRequest":
        if self.new_password != self.new_password_confirm:
            raise ValueError("새 비밀번호가 일치하지 않습니다.")
        if self.current_password == self.new_password:
            raise ValueError("새 비밀번호는 현재 비밀번호와 달라야 합니다.")
        return self


class MessageResponse(BaseModel):
    message: str


class RegisterResponse(BaseModel):
    message: str
    status: str


class AccountStatusResponse(BaseModel):
    status: Literal["none", "pending", "active", "rejected"]


class PendingUserItem(BaseModel):
    id: UUID
    email: str
    status: str
    email_verified: bool
    created_at: datetime


class PendingUsersResponse(BaseModel):
    items: list[PendingUserItem]

__all__ = [
    "AccountStatusResponse",
    "ChangePasswordRequest",
    "LoginRequest",
    "MeResponse",
    "MessageResponse",
    "PendingUserItem",
    "PendingUsersResponse",
    "RegisterRequest",
    "RegisterResponse",
    "SignupEmailSendRequest",
    "SignupEmailSendResponse",
    "SignupTokenCheckResponse",
    "TokenResponse",
]
