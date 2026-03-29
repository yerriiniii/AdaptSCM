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

    @property
    def database_enabled(self) -> bool:
        return bool(self.database_url)

    @property
    def s3_enabled(self) -> bool:
        return bool(self.s3_bucket and self.aws_access_key_id and self.aws_secret_access_key)


def get_runtime_settings() -> RuntimeSettings:
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
    )
