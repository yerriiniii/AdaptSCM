import uuid
from collections.abc import Generator
from datetime import datetime, timezone

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.engine import Engine
from sqlalchemy.orm import Session, sessionmaker

from app.shared.config import get_runtime_settings
from app.shared.db.base import Base


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)

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
        "item": {
            "brand": "VARCHAR(255)",
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
        if {"product_groups", "product_locales", "item", "item_mapping"}.issubset(existing_tables):
            item_count = connection.execute(text("SELECT COUNT(*) FROM item")).scalar() or 0
            item_mapping_count = connection.execute(text("SELECT COUNT(*) FROM item_mapping")).scalar() or 0
            if int(item_count) == 0 and int(item_mapping_count) == 0:
                legacy_groups = connection.execute(
                    text(
                        """
                        SELECT id, representative_name, upload_updated_at, manual_updated_at, updated_at
                        FROM product_groups
                        ORDER BY representative_name ASC
                        """
                    )
                ).mappings().all()
                legacy_locales = connection.execute(
                    text(
                        """
                        SELECT id, product_group_id, country_code, name, sku, updated_at
                        FROM product_locales
                        ORDER BY product_group_id ASC, country_code ASC
                        """
                    )
                ).mappings().all()

                for row in legacy_groups:
                    connection.execute(
                        text(
                            """
                            INSERT INTO item (id, kr_name, upload_updated_at, manual_updated_at, updated_at)
                            VALUES (:id, :kr_name, :upload_updated_at, :manual_updated_at, :updated_at)
                            """
                        ),
                        {
                            "id": row.get("id"),
                            "kr_name": row.get("representative_name"),
                            "upload_updated_at": row.get("upload_updated_at"),
                            "manual_updated_at": row.get("manual_updated_at"),
                            "updated_at": row.get("updated_at"),
                        },
                    )

                for row in legacy_locales:
                    connection.execute(
                        text(
                            """
                            INSERT INTO item_mapping (id, item_id, country_code, name, sku, updated_at)
                            VALUES (:id, :item_id, :country_code, :name, :sku, :updated_at)
                            """
                        ),
                        {
                            "id": row.get("id"),
                            "item_id": row.get("product_group_id"),
                            "country_code": row.get("country_code"),
                            "name": row.get("name"),
                            "sku": row.get("sku"),
                            "updated_at": row.get("updated_at"),
                        },
                    )

        if {"sku", "item", "item_mapping"}.issubset(existing_tables):
            item_count = connection.execute(text("SELECT COUNT(*) FROM item")).scalar() or 0
            item_mapping_count = connection.execute(text("SELECT COUNT(*) FROM item_mapping")).scalar() or 0
            if int(item_count) == 0 and int(item_mapping_count) == 0:
                sku_columns = {column["name"] for column in inspector.get_columns("sku")}
                select_columns = ["description", "kr", "us", "tw", "vn", "sg", "au", "uk", "ae"]
                optional_columns = ["updated_at", "upload_updated_at", "manual_updated_at"]
                available_optional = [column for column in optional_columns if column in sku_columns]
                selected = ", ".join(select_columns + available_optional)
                legacy_rows = connection.execute(text(f"SELECT {selected} FROM sku ORDER BY description ASC")).mappings().all()
                for row in legacy_rows:
                    kr_name = str(row.get("description") or "").strip()
                    if not kr_name:
                        continue
                    item_id = str(uuid.uuid4())
                    updated_at = row.get("updated_at") or _utc_now()
                    upload_updated_at = row.get("upload_updated_at")
                    manual_updated_at = row.get("manual_updated_at")
                    if upload_updated_at is None and manual_updated_at is None:
                        upload_updated_at = updated_at

                    connection.execute(
                        text(
                            """
                            INSERT INTO item (
                                id,
                                kr_name,
                                upload_updated_at,
                                manual_updated_at,
                                updated_at
                            ) VALUES (
                                :id,
                                :kr_name,
                                :upload_updated_at,
                                :manual_updated_at,
                                :updated_at
                            )
                            """
                        ),
                        {
                            "id": item_id,
                            "kr_name": kr_name,
                            "upload_updated_at": upload_updated_at,
                            "manual_updated_at": manual_updated_at,
                            "updated_at": updated_at,
                        },
                    )

                    for country_code, column_name in [
                        ("KR", "kr"),
                        ("US", "us"),
                        ("TW", "tw"),
                        ("VN", "vn"),
                        ("SG", "sg"),
                        ("AU", "au"),
                        ("UK", "uk"),
                        ("AE", "ae"),
                    ]:
                        sku = str(row.get(column_name) or "").strip()
                        if not sku:
                            continue
                        connection.execute(
                            text(
                                """
                                INSERT INTO item_mapping (
                                    id,
                                    item_id,
                                    country_code,
                                    name,
                                    sku,
                                    updated_at
                                ) VALUES (
                                    :id,
                                    :item_id,
                                    :country_code,
                                    :name,
                                    :sku,
                                    :updated_at
                                )
                                """
                            ),
                            {
                                "id": str(uuid.uuid4()),
                                "item_id": item_id,
                                "country_code": country_code,
                                "name": kr_name,
                                "sku": sku,
                                "updated_at": updated_at,
                            },
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


def _ensure_item_mapping_allows_multiple_locales_per_country(engine: Engine) -> None:
    """(item_id, country_code) 유니크 제거 → 동일 item 에 TW 등 국가별 SKU 여러 행 허용."""
    inspector = inspect(engine)
    if "item_mapping" not in inspector.get_table_names():
        return
    dialect = engine.dialect.name
    with engine.begin() as conn:
        if dialect == "sqlite":
            conn.execute(text("DROP INDEX IF EXISTS ix_item_mapping_item_country"))
        elif dialect == "postgresql":
            conn.execute(text("DROP INDEX IF EXISTS ix_item_mapping_item_country"))
        elif dialect == "mysql":
            try:
                conn.execute(text("ALTER TABLE item_mapping DROP INDEX ix_item_mapping_item_country"))
            except Exception:
                pass
        if dialect in ("sqlite", "postgresql", "mysql"):
            if dialect == "mysql":
                try:
                    conn.execute(
                        text(
                            "CREATE INDEX ix_item_mapping_item_country ON item_mapping (item_id, country_code)"
                        )
                    )
                except Exception:
                    pass
            else:
                conn.execute(
                    text(
                        "CREATE INDEX IF NOT EXISTS ix_item_mapping_item_country "
                        "ON item_mapping (item_id, country_code)"
                    )
                )


def initialize_database() -> None:
    engine = get_engine()
    Base.metadata.create_all(bind=engine)
    _ensure_inventory_columns(engine)
    _ensure_item_mapping_allows_multiple_locales_per_country(engine)


def test_database_connection() -> bool:
    engine = get_engine()
    with engine.connect() as connection:
        connection.execute(text("SELECT 1"))
    return True
