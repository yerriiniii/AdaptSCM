from collections.abc import Generator

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session, sessionmaker

from app.shared.config import get_runtime_settings
from app.shared.db.base import Base

_engine: Engine | None = None
_session_factory: sessionmaker[Session] | None = None


def _ensure_inventory_columns(engine: Engine) -> None:
    inspector = inspect(engine)
    existing_tables = set(inspector.get_table_names())
    column_specs = {
        "inventory_rows": {
            "sku": "VARCHAR(255)",
            "warehouse": "VARCHAR(255)",
        },
        "inventory_aggregates": {
            "sku": "VARCHAR(255)",
            "warehouse": "VARCHAR(255)",
        },
        "sku": {
            "updated_at": "TIMESTAMP WITH TIME ZONE DEFAULT NOW() NOT NULL",
            "upload_updated_at": "TIMESTAMP WITH TIME ZONE",
            "manual_updated_at": "TIMESTAMP WITH TIME ZONE",
        },
    }

    with engine.begin() as connection:
        for table_name, columns in column_specs.items():
            if table_name not in existing_tables:
                continue
            existing_columns = {column["name"] for column in inspector.get_columns(table_name)}
            for column_name, column_type in columns.items():
                if column_name in existing_columns:
                    continue
                connection.execute(text(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {column_type}"))

        if "inventory_rows" in existing_tables:
            row_columns = {column["name"] for column in inspector.get_columns("inventory_rows")}
            row_sources = [column for column in ["sku", "raw_item_code", "item_code"] if column in row_columns]
            if row_sources:
                row_expr = ", ".join(row_sources)
                connection.execute(
                    text(
                        f"""
                        UPDATE inventory_rows
                        SET sku = COALESCE({row_expr})
                        WHERE sku IS NULL
                        """
                    )
                )
        if "inventory_aggregates" in existing_tables:
            aggregate_columns = {column["name"] for column in inspector.get_columns("inventory_aggregates")}
            aggregate_sources = [column for column in ["sku", "raw_item_code", "item_code"] if column in aggregate_columns]
            if aggregate_sources:
                aggregate_expr = ", ".join(aggregate_sources)
                connection.execute(
                    text(
                        f"""
                        UPDATE inventory_aggregates
                        SET sku = COALESCE({aggregate_expr})
                        WHERE sku IS NULL
                        """
                    )
                )
        if "sku" in existing_tables:
            sku_columns = {column["name"] for column in inspector.get_columns("sku")}
            if {"updated_at", "upload_updated_at", "manual_updated_at"}.issubset(sku_columns):
                connection.execute(
                    text(
                        """
                        UPDATE sku
                        SET upload_updated_at = updated_at
                        WHERE upload_updated_at IS NULL AND manual_updated_at IS NULL
                        """
                    )
                )


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
    _ensure_inventory_columns(engine)


def test_database_connection() -> bool:
    engine = get_engine()
    with engine.connect() as connection:
        connection.execute(text("SELECT 1"))
    return True
