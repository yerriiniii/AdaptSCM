import secrets

from fastapi import HTTPException, status

from shared.config import get_runtime_settings


def verify_admin_access_code(code: str) -> None:
    settings = get_runtime_settings()
    expected = settings.admin_access_code
    if not expected:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="관리자 인증번호(ADMIN_ACCESS_CODE)가 서버에 설정되지 않았습니다.",
        )
    provided = str(code or "").strip()
    if not provided:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="인증번호를 입력해 주세요.")
    if not secrets.compare_digest(provided, expected):
        raise HTTPException(status_code=status.HTTP_401_UNAUTHORIZED, detail="인증번호가 올바르지 않습니다.")


__all__ = ["verify_admin_access_code"]
