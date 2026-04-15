import logging
import os
import threading
import time
import uuid
from collections.abc import Generator
from datetime import datetime, timezone

from sqlalchemy import create_engine, inspect, text
from sqlalchemy.engine import Engine
from sqlalchemy.exc import OperationalError
from sqlalchemy.orm import Session, sessionmaker

from app.shared.config import get_runtime_settings
from app.shared.db.base import Base

_log = logging.getLogger(__name__)

_ITEM_MAPPING_COUNTRY_INDEX = "ix_item_mapping_item_country"


def _utc_now() -> datetime:
    return datetime.now(timezone.utc)

_engine: Engine | None = None
_session_factory: sessionmaker[Session] | None = None
_engine_creation_lock = threading.Lock()
_schema_init_lock = threading.Lock()
_schema_initialized = False


def _engine_connect_kwargs(database_url: str) -> dict:
    """느리거나 끊긴 DB에 TCP가 무한 대기하지 않도록 타임아웃(초)."""
    url = str(database_url or "").lower()
    if "postgresql" not in url:
        return {}
    try:
        sec = max(1, min(120, int(os.getenv("DB_CONNECT_TIMEOUT", "10"))))
    except ValueError:
        sec = 10
    # psycopg (v3): connect_timeout
    return {"connect_args": {"connect_timeout": sec}}


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
            "barcode": "VARCHAR(255)",
        },
        "purchase_inbound_lines": {
            "actual_inbound_note": "VARCHAR(128)",
            "delivery_available_date": "DATE",
            "expected_inbound_date": "DATE",
            "line_memo": "TEXT",
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
    """엔진·세션 팩토리만 생성. 스키마는 `initialize_database()`에서 단일 경로로 적용."""
    global _engine
    global _session_factory

    if _engine is not None:
        return _engine

    with _engine_creation_lock:
        if _engine is not None:
            return _engine

        settings = get_runtime_settings()
        if not settings.database_url:
            raise RuntimeError("DATABASE_URL 환경변수가 설정되지 않았습니다.")

        extra = _engine_connect_kwargs(settings.database_url)
        _engine = create_engine(
            settings.database_url,
            echo=settings.db_echo,
            pool_pre_ping=True,
            **extra,
        )
        _session_factory = sessionmaker(bind=_engine, autoflush=False, autocommit=False)
        return _engine


def get_db_session() -> Generator[Session, None, None]:
    global _session_factory

    if is_database_configured():
        initialize_database()

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

    # PostgreSQL: 트랜잭션 안 DROP INDEX + 풀러 statement_timeout 에서 기동 실패할 수 있음.
    # autocommit + statement_timeout 해제 후 CONCURRENTLY(가능 시)로 처리.
    if dialect == "postgresql":
        index_names = {ix.get("name") for ix in inspector.get_indexes("item_mapping")}
        if _ITEM_MAPPING_COUNTRY_INDEX not in index_names:
            return
        try:
            with engine.connect().execution_options(isolation_level="AUTOCOMMIT") as conn:
                # 무제한 대기(기존 statement_timeout=0)는 --reload 직후 등 락 경합 시 영구 정지처럼 보일 수 있음
                try:
                    lock_s = max(5, min(300, int(os.getenv("PG_INDEX_MIGRATION_LOCK_TIMEOUT_SEC", "60"))))
                except ValueError:
                    lock_s = 60
                try:
                    stmt_s = max(30, min(3600, int(os.getenv("PG_INDEX_MIGRATION_STATEMENT_TIMEOUT_SEC", "600"))))
                except ValueError:
                    stmt_s = 600
                conn.execute(text(f"SET lock_timeout TO '{lock_s}s'"))
                conn.execute(text(f"SET statement_timeout TO '{stmt_s}s'"))
                try:
                    conn.execute(
                        text(f"DROP INDEX CONCURRENTLY IF EXISTS {_ITEM_MAPPING_COUNTRY_INDEX}")
                    )
                except OperationalError as exc:
                    _log.warning(
                        "DROP INDEX CONCURRENTLY %s failed (%s); retrying without CONCURRENTLY",
                        _ITEM_MAPPING_COUNTRY_INDEX,
                        exc,
                    )
                    conn.execute(text(f"DROP INDEX IF EXISTS {_ITEM_MAPPING_COUNTRY_INDEX}"))
                try:
                    conn.execute(
                        text(
                            f"CREATE INDEX CONCURRENTLY IF NOT EXISTS {_ITEM_MAPPING_COUNTRY_INDEX} "
                            "ON item_mapping (item_id, country_code)"
                        )
                    )
                except OperationalError as exc:
                    _log.warning(
                        "CREATE INDEX CONCURRENTLY %s failed (%s); retrying regular CREATE INDEX",
                        _ITEM_MAPPING_COUNTRY_INDEX,
                        exc,
                    )
                    conn.execute(
                        text(
                            f"CREATE INDEX IF NOT EXISTS {_ITEM_MAPPING_COUNTRY_INDEX} "
                            "ON item_mapping (item_id, country_code)"
                        )
                    )
        except OperationalError as exc:
            _log.error(
                "item_mapping 인덱스 조정 실패(앱은 계속 기동됨): %s. "
                "필요 시 수동 실행: DROP INDEX CONCURRENTLY IF EXISTS %s; "
                "CREATE INDEX CONCURRENTLY IF NOT EXISTS %s ON item_mapping (item_id, country_code);",
                exc,
                _ITEM_MAPPING_COUNTRY_INDEX,
                _ITEM_MAPPING_COUNTRY_INDEX,
            )
        return

    with engine.begin() as conn:
        if dialect == "sqlite":
            conn.execute(text(f"DROP INDEX IF EXISTS {_ITEM_MAPPING_COUNTRY_INDEX}"))
        elif dialect == "mysql":
            try:
                conn.execute(text(f"ALTER TABLE item_mapping DROP INDEX {_ITEM_MAPPING_COUNTRY_INDEX}"))
            except Exception:
                pass
        if dialect in ("sqlite", "mysql"):
            if dialect == "mysql":
                try:
                    conn.execute(
                        text(
                            f"CREATE INDEX {_ITEM_MAPPING_COUNTRY_INDEX} "
                            "ON item_mapping (item_id, country_code)"
                        )
                    )
                except Exception:
                    pass
            else:
                conn.execute(
                    text(
                        f"CREATE INDEX IF NOT EXISTS {_ITEM_MAPPING_COUNTRY_INDEX} "
                        "ON item_mapping (item_id, country_code)"
                    )
                )


def _ensure_purchase_orders_schema(engine: Engine) -> None:
    """발주 예정(날짜 없음) 대응: order_date_note 추가, order_date NULL 허용(PostgreSQL/MySQL)."""
    inspector = inspect(engine)
    if "purchase_orders" not in inspector.get_table_names():
        return
    dialect = engine.dialect.name
    with engine.begin() as conn:
        cols = {c["name"]: c for c in inspector.get_columns("purchase_orders")}
        if "order_date_note" not in cols:
            conn.execute(text("ALTER TABLE purchase_orders ADD COLUMN order_date_note VARCHAR(128)"))
        cols = {c["name"]: c for c in inspect(engine).get_columns("purchase_orders")}
        od = cols.get("order_date")
        if od and od.get("nullable") is False:
            if dialect == "postgresql":
                conn.execute(text("ALTER TABLE purchase_orders ALTER COLUMN order_date DROP NOT NULL"))
            elif dialect == "mysql":
                conn.execute(text("ALTER TABLE purchase_orders MODIFY COLUMN order_date DATE NULL"))


_OLD_PO_UQ_ERP_ONLY = "uq_purchase_orders_erp_po_number"
_NEW_PO_UQ_COMPOSITE = "uq_purchase_orders_erp_sku_order_date"


def _ensure_purchase_orders_erp_sku_order_date_unique(engine: Engine) -> None:
    """ERP PO 단독 유니크 → (ERP PO, SKU, 발주일) 복합 유니크로 전환."""
    inspector = inspect(engine)
    if "purchase_orders" not in inspector.get_table_names():
        return
    dialect = engine.dialect.name

    def _unique_constraint_names() -> set[str]:
        ins = inspect(engine)
        return {uc["name"] for uc in ins.get_unique_constraints("purchase_orders") if uc.get("name")}

    def _has_new_composite(ins: object) -> bool:
        for uc in ins.get_unique_constraints("purchase_orders"):
            cols = tuple(uc.get("column_names") or ())
            if cols == ("erp_po_number", "sku", "order_date"):
                return True
        for ix in ins.get_indexes("purchase_orders"):
            if not ix.get("unique"):
                continue
            cols = tuple(ix.get("column_names") or ())
            if cols == ("erp_po_number", "sku", "order_date"):
                return True
        return False

    names = _unique_constraint_names()
    with engine.begin() as conn:
        if _OLD_PO_UQ_ERP_ONLY in names:
            if dialect == "postgresql":
                conn.execute(
                    text(f'ALTER TABLE purchase_orders DROP CONSTRAINT IF EXISTS "{_OLD_PO_UQ_ERP_ONLY}"')
                )
            elif dialect == "mysql":
                conn.execute(text(f"ALTER TABLE purchase_orders DROP INDEX {_OLD_PO_UQ_ERP_ONLY}"))
            elif dialect == "sqlite":
                try:
                    conn.execute(
                        text(f"ALTER TABLE purchase_orders DROP CONSTRAINT {_OLD_PO_UQ_ERP_ONLY}")
                    )
                except Exception:
                    conn.execute(text(f"DROP INDEX IF EXISTS {_OLD_PO_UQ_ERP_ONLY}"))

    ins2 = inspect(engine)
    if _has_new_composite(ins2):
        return

    with engine.begin() as conn:
        if dialect == "postgresql":
            conn.execute(
                text(
                    f'ALTER TABLE purchase_orders ADD CONSTRAINT "{_NEW_PO_UQ_COMPOSITE}" '
                    "UNIQUE (erp_po_number, sku, order_date)"
                )
            )
        elif dialect == "mysql":
            conn.execute(
                text(
                    f"ALTER TABLE purchase_orders ADD UNIQUE KEY {_NEW_PO_UQ_COMPOSITE} "
                    "(erp_po_number, sku, order_date)"
                )
            )
        elif dialect == "sqlite":
            conn.execute(
                text(
                    f"CREATE UNIQUE INDEX IF NOT EXISTS {_NEW_PO_UQ_COMPOSITE} "
                    "ON purchase_orders (erp_po_number, sku, order_date)"
                )
            )


def _run_schema_init_if_needed(engine: Engine) -> None:
    """테이블 생성·경량 마이그레이션. 프로세스당 1회."""
    global _schema_initialized
    with _schema_init_lock:
        if _schema_initialized:
            return
        t0 = time.perf_counter()
        _log.info("DB 스키마 초기화 시작…")
        Base.metadata.create_all(bind=engine)
        _log.info("create_all 완료 (%.2fs)", time.perf_counter() - t0)
        t1 = time.perf_counter()
        _ensure_inventory_columns(engine)
        _log.info("inventory 컬럼 보정 완료 (%.2fs)", time.perf_counter() - t1)
        t2 = time.perf_counter()
        _ensure_item_mapping_allows_multiple_locales_per_country(engine)
        _log.info("item_mapping 인덱스 보정 완료 (%.2fs)", time.perf_counter() - t2)
        t3 = time.perf_counter()
        _ensure_purchase_orders_schema(engine)
        _log.info("purchase_orders 스키마 보정 완료 (%.2fs)", time.perf_counter() - t3)
        t4 = time.perf_counter()
        _ensure_purchase_orders_erp_sku_order_date_unique(engine)
        _log.info("purchase_orders 유니크 보정 완료 (%.2fs)", time.perf_counter() - t4)
        _schema_initialized = True
        _log.info("DB 스키마 초기화 전체 완료 (총 %.2fs)", time.perf_counter() - t0)


def initialize_database() -> None:
    """테이블·경량 마이그레이션 1회. 스크립트·startup·첫 세션에서 호출 가능."""
    if not is_database_configured():
        return
    _run_schema_init_if_needed(get_engine())


def test_database_connection() -> bool:
    if not is_database_configured():
        raise RuntimeError("DATABASE_URL 환경변수가 설정되지 않았습니다.")
    initialize_database()
    engine = get_engine()
    with engine.connect() as connection:
        connection.execute(text("SELECT 1"))
    return True
