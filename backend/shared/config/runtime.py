import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(dotenv_path=Path(__file__).resolve().parents[3] / ".env")


def _read_bool(value: str | None, default: bool = False) -> bool:
    if value is None:
        return default
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def _read_csv(value: str | None, default: list[str] | None = None) -> list[str]:
    if value is None:
        return list(default or [])
    return [item.strip() for item in str(value).split(",") if item.strip()]


@dataclass(frozen=True)
class RuntimeSettings:
    database_url: str | None
    aws_access_key_id: str | None
    aws_secret_access_key: str | None
    aws_region: str
    s3_bucket: str | None
    s3_prefix: str
    db_echo: bool
    frontend_origins: list[str]
    # JWT: 서명에 사용. 운영에서는 긴 무작위 문자열 필수.
    jwt_secret: str
    jwt_algorithm: str
    jwt_expire_minutes: int
    # 최초 기동 시 1회 생성(이메일이 없을 때만). 이후엔 env 제거·변경해도 기존 계정은 유지.
    initial_admin_email: str | None
    initial_admin_password: str | None
    # 가입 시 인증 링크에 쓰는 프론트 URL (끝 슬래시 없이)
    public_app_url: str
    # 가입 환영 메일 (SMTP 없으면 링크만 서버 로그에 남음)
    email_smtp_host: str | None
    email_smtp_port: int
    email_smtp_user: str | None
    email_smtp_password: str | None
    email_smtp_from: str | None
    email_smtp_use_tls: bool
    # 이메일 인증 토큰 유효 시간(초)
    email_verification_ttl_sec: int

    @property
    def database_enabled(self) -> bool:
        return bool(self.database_url)

    @property
    def s3_enabled(self) -> bool:
        return bool(self.s3_bucket and self.aws_access_key_id and self.aws_secret_access_key)


def get_runtime_settings() -> RuntimeSettings:
    def _read_int(value: str | None, default: int) -> int:
        if value is None or not str(value).strip():
            return default
        try:
            return int(value)
        except ValueError:
            return default

    return RuntimeSettings(
        database_url=os.getenv("DATABASE_URL"),
        aws_access_key_id=os.getenv("AWS_ACCESS_KEY_ID"),
        aws_secret_access_key=os.getenv("AWS_SECRET_ACCESS_KEY"),
        aws_region=os.getenv("AWS_REGION", "ap-northeast-2"),
        s3_bucket=os.getenv("S3_BUCKET"),
        s3_prefix=os.getenv("S3_PREFIX", "inventory"),
        db_echo=_read_bool(os.getenv("DB_ECHO"), default=False),
        frontend_origins=_read_csv(
            os.getenv("FRONTEND_ORIGINS"),
            default=["http://localhost:5173", "http://127.0.0.1:5173"],
        ),
        jwt_secret=(os.getenv("JWT_SECRET") or "dev-only-change-me").strip(),
        jwt_algorithm=os.getenv("JWT_ALGORITHM", "HS256").strip() or "HS256",
        jwt_expire_minutes=max(1, _read_int(os.getenv("JWT_EXPIRE_MINUTES"), 60 * 8)),
        initial_admin_email=(os.getenv("INITIAL_ADMIN_EMAIL") or "").strip() or None,
        initial_admin_password=os.getenv("INITIAL_ADMIN_PASSWORD") or None,
        public_app_url=(os.getenv("PUBLIC_APP_URL") or "http://localhost:5173").rstrip("/"),
        email_smtp_host=(os.getenv("EMAIL_SMTP_HOST") or "").strip() or None,
        email_smtp_port=max(1, _read_int(os.getenv("EMAIL_SMTP_PORT"), 587)),
        email_smtp_user=(os.getenv("EMAIL_SMTP_USER") or "").strip() or None,
        email_smtp_password=os.getenv("EMAIL_SMTP_PASSWORD") or None,
        email_smtp_from=(os.getenv("EMAIL_SMTP_FROM") or "").strip() or None,
        email_smtp_use_tls=_read_bool(os.getenv("EMAIL_SMTP_USE_TLS", "true"), default=True),
        # 가입용 이메일 인증 링크(토큰) 유효 시간. 기본 5분(300초). env 로 조정(최소 60초).
        email_verification_ttl_sec=max(60, _read_int(os.getenv("EMAIL_VERIFICATION_TTL_SEC"), 300)),
    )
