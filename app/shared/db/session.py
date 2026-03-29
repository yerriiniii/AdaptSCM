from collections.abc import Generator

from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session, sessionmaker

from app.shared.config import get_runtime_settings
from app.shared.db.base import Base

_engine: Engine | None = None
_session_factory: sessionmaker[Session] | None = None


def is_database_configured() -> bool:
    settings = get_runtime_settings()
    return settings.database_enabled


def get_engine() -> Engine:
    global _engine
    global _session_factory

    if _engine is not None:
        return _engine

    settings = get_runtime_settings()
    if not settings.database_url:
        raise RuntimeError("DATABASE_URL 환경변수가 설정되지 않았습니다.")

    _engine = create_engine(
        settings.database_url,
        echo=settings.db_echo,
        pool_pre_ping=True,
    )
    _session_factory = sessionmaker(bind=_engine, autoflush=False, autocommit=False)
    return _engine


def get_db_session() -> Generator[Session, None, None]:
    global _session_factory

    if _session_factory is None:
        get_engine()

    if _session_factory is None:
        raise RuntimeError("DB 세션 팩토리를 초기화하지 못했습니다.")

    session = _session_factory()
    try:
        yield session
    finally:
        session.close()


def initialize_database() -> None:
    engine = get_engine()
    Base.metadata.create_all(bind=engine)


def test_database_connection() -> bool:
    engine = get_engine()
    with engine.connect() as connection:
        connection.execute(text("SELECT 1"))
    return True
