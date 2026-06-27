from datetime import datetime
from uuid import UUID

from pydantic import BaseModel, Field


class AdminAccessCodeRequest(BaseModel):
    code: str = Field(..., min_length=1, max_length=64)


class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    expires_in: int  # seconds


class MeResponse(BaseModel):
    id: UUID
    email: str
    status: str
    is_admin: bool
    created_at: datetime


__all__ = [
    "AdminAccessCodeRequest",
    "MeResponse",
    "TokenResponse",
]
