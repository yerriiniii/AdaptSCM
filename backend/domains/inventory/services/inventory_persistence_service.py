import logging
import os
import uuid
from collections.abc import Iterable
from datetime import date

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import and_, delete, func, or_, select
from sqlalchemy.orm import Session

from domains.inventory.models import InventoryAggregate, InventoryRow, UploadedFile
from domains.inventory.services.inventory_aggregate_service import (
    _detect_country,
    _normalize_country_code,
    _normalize_item_code,
    _normalize_level_value,
    _parse_date_from_filename,
    _parse_date_series,
    _read_inventory_bytes,
    _read_inventory_file,
)
from domains.inventory.services.inventory_item_mapping_resolve import (
    apply_db_locale_columns_to_inventory_merged,
    apply_db_locale_columns_to_raw_inventory_df,
    attach_item_mapping_kr_to_inventory_rows,
    validate_inventory_row_frame_item_mapping,
)
from shared.config import get_runtime_settings
from shared.storage import create_presigned_upload_url, get_s3_client, is_s3_configured

logger = logging.getLogger(__name__)


def _read_upload_bytes(upload_file: UploadFile) -> bytes:
    upload_file.file.seek(0)
    raw = upload_file.file.read()
    upload_file.file.seek(0)
    return raw


def _resolve_country_code(
    df: pd.DataFrame,
    upload_file: UploadFile,
    idx: int,
    file_countries: list[str] | None,
) -> str:
    override_country = None
    if file_countries and idx < len(file_countries):
        override_country = _normalize_country_code(file_countries[idx])
        if file_countries[idx] and override_country is None:
            raise HTTPException(
                status_code=400,
                detail=f"{upload_file.filename}: 국가 코드는 영문 2~4자리여야 합니다.",
            )

    if override_country is not None:
        return override_country

    warehouse_col = df["warehouse"] if "warehouse" in df.columns else ""
    if isinstance(warehouse_col, str):
        return _detect_country(warehouse_col, upload_file.filename)

    detected = _detect_country(
        str(warehouse_col.fillna("").astype(str).iloc[0]) if len(df) else "",
        upload_file.filename,
    )
    return detected or "KR"


def _resolve_base_date(
    df: pd.DataFrame,
    upload_file: UploadFile,
    idx: int,
    file_dates: list[str] | None,
    country_code: str,
) -> date | None:
    filename_date = _parse_date_from_filename(upload_file.filename)

    if file_dates and idx < len(file_dates):
        raw_override = str(file_dates[idx] or "").strip()
        if raw_override:
            parsed_override = pd.to_datetime(raw_override, errors="coerce")
            if pd.isna(parsed_override):
                raise HTTPException(
                    status_code=400,
                    detail=f"{upload_file.filename}: 설정 탭의 날짜 형식이 올바르지 않습니다.",
                )
            return parsed_override.date()

    is_kr_file = country_code == "KR"
    if is_kr_file and filename_date is not None:
        return filename_date

    if "date" in df.columns:
        parsed_date = _parse_date_series(df["date"])
        valid = parsed_date.dropna()
        if not valid.empty:
            representative = pd.Series(valid).value_counts().index[0]
            if isinstance(representative, pd.Timestamp):
                return representative.date()
            return representative

    if filename_date is not None:
        return filename_date

    if is_kr_file:
        return pd.Timestamp.today().date()

    return None


def _country_type(country_code: str) -> str:
    return "KR" if country_code == "KR" else "OVERSEAS"


def _build_s3_key(prefix: str, country_type: str, country_code: str, base_date: date | None, stored_name: str) -> str:
    parts = [
        prefix.strip("/"),
        country_type.lower(),
        country_code.lower(),
        base_date.isoformat() if base_date else "unknown-date",
        stored_name,
    ]
    return "/".join(part for part in parts if part)


def _parse_optional_date(raw_value: str | None, filename: str) -> date | None:
    raw = str(raw_value or "").strip()
    if not raw:
        return None
    parsed = pd.to_datetime(raw, errors="coerce")
    if pd.isna(parsed):
        raise HTTPException(status_code=400, detail=f"{filename}: 날짜 형식이 올바르지 않습니다.")
    return parsed.date()


def _nullable_string(value: object) -> str | None:
    if pd.isna(value):
        return None

    raw = str(value).strip()
    return raw or None


def _next_aggregation_version(db: Session) -> int:
    current = db.execute(select(func.max(InventoryAggregate.aggregation_version))).scalar_one_or_none()
    return int(current or 0) + 1


def _serialize_uploaded_file(uploaded_file: UploadedFile, client_id: str | None = None) -> dict:
    fd = getattr(uploaded_file, "file_domain", None) or "inventory"
    return {
        "client_id": client_id,
        "file_id": str(uploaded_file.id),
        "name": uploaded_file.original_name,
        "country": uploaded_file.country_code,
        "date": uploaded_file.base_date.isoformat() if uploaded_file.base_date else "",
        "size": uploaded_file.size or 0,
        "status": uploaded_file.status,
        "file_domain": fd,
    }


def _scope_filter(country_code: str, base_date: date | None):
    inv_only = or_(
        InventoryAggregate.source_domain == "INVENTORY",
        InventoryAggregate.source_domain.is_(None),
    )
    if base_date is None:
        return and_(
            InventoryAggregate.country_code == country_code,
            InventoryAggregate.base_date.is_(None),
            inv_only,
        )
    return and_(
        InventoryAggregate.country_code == country_code,
        InventoryAggregate.base_date == base_date,
        inv_only,
    )


def _uploaded_file_scope_filter(country_code: str, base_date: date | None):
    if base_date is None:
        return and_(UploadedFile.country_code == country_code, UploadedFile.base_date.is_(None))
    return and_(UploadedFile.country_code == country_code, UploadedFile.base_date == base_date)


def _resolved_upload_s3_bucket_and_key(uploaded_file: UploadedFile) -> tuple[str | None, str | None]:
    """DB에 s3_key가 비어 있어도 stored_name·국가·날짜로 키를 재구성해 삭제를 시도한다."""
    settings = get_runtime_settings()
    bucket = (uploaded_file.s3_bucket or "").strip() or (settings.s3_bucket or "")
    key = (uploaded_file.s3_key or "").strip()
    stored = (uploaded_file.stored_name or "").strip()
    if not key and stored:
        key = _build_s3_key(
            prefix=settings.s3_prefix,
            country_type=_country_type(uploaded_file.country_code),
            country_code=uploaded_file.country_code,
            base_date=uploaded_file.base_date,
            stored_name=stored,
        )
    return (bucket or None, key or None)


def _delete_s3_object(bucket: str | None, key: str | None) -> None:
    if not is_s3_configured():
        logger.warning("S3 미설정으로 객체 삭제를 건너뜁니다 (bucket=%r, key=%r).", bucket, key)
        return
    if not bucket or not key:
        logger.warning("S3 bucket 또는 key가 없어 삭제를 건너뜁니다 (bucket=%r, key=%r).", bucket, key)
        return
    client = get_s3_client()
    client.delete_object(Bucket=bucket, Key=key)


def _s3_prefix_for_country_scope(settings, country_code: str) -> str:
    """해당 국가 재고 파일이 두는 공통 접두사(하위 날짜 폴더까지 스윕)."""
    root = (settings.s3_prefix or "inventory").strip().strip("/")
    if not root:
        return ""
    cc = str(country_code or "").strip().upper()
    if not cc:
        return ""
    if cc == "SHIPMENT":
        return f"{root}/shipment/"
    ct = "kr" if cc == "KR" else "overseas"
    return f"{root}/{ct}/{cc.lower()}/"


def _s3_prefix_inventory_root(settings) -> str:
    root = (settings.s3_prefix or "inventory").strip().strip("/")
    if not root:
        return ""
    return f"{root}/"


def _delete_s3_prefix(bucket: str | None, prefix: str) -> int:
    """prefix 아래 객체를 모두 삭제한다. 삭제한 객체 수를 반환한다."""
    if not bucket or not prefix or not is_s3_configured():
        return 0
    client = get_s3_client()
    deleted = 0
    paginator = client.get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        contents = page.get("Contents") or []
        if not contents:
            continue
        keys = [{"Key": c["Key"]} for c in contents]
        for i in range(0, len(keys), 1000):
            batch = keys[i : i + 1000]
            client.delete_objects(Bucket=bucket, Delete={"Objects": batch})
            deleted += len(batch)
    return deleted


def _rebuild_aggregate_scope(db: Session, country_code: str, base_date: date | None) -> None:
    db.execute(delete(InventoryAggregate).where(_scope_filter(country_code, base_date)))

    uploaded_file_ids = db.execute(
        select(UploadedFile.id).where(_uploaded_file_scope_filter(country_code, base_date))
    ).scalars().all()
    if not uploaded_file_ids:
        return

    rows = db.execute(
        select(
            InventoryRow.sku,
            InventoryRow.description,
            InventoryRow.supplier,
            InventoryRow.level,
            InventoryRow.warehouse,
            InventoryRow.quantity,
        ).where(InventoryRow.uploaded_file_id.in_(uploaded_file_ids))
    ).all()
    if not rows:
        return

    frame = pd.DataFrame(
        [
            {
                "sku": row.sku,
                "description": row.description,
                "supplier": row.supplier,
                "level": row.level,
                "warehouse": row.warehouse,
                "quantity": int(row.quantity or 0),
            }
            for row in rows
        ]
    )
    grouped = (
        frame.groupby(
            ["sku", "description", "supplier", "level", "warehouse"],
            dropna=False,
            as_index=False,
        )["quantity"]
        .sum()
    )
    aggregation_version = _next_aggregation_version(db)
    for record in grouped.to_dict(orient="records"):
        db.add(
            InventoryAggregate(
                country_code=country_code,
                base_date=base_date,
                sku=str(record["sku"]),
                description=_nullable_string(record["description"]),
                supplier=_nullable_string(record["supplier"]),
                level=_nullable_string(record["level"]),
                warehouse=_nullable_string(record["warehouse"]),
                total_quantity=int(record["quantity"]),
                aggregation_version=aggregation_version,
                source_domain="INVENTORY",
            )
        )


def _rebuild_aggregate_scopes(db: Session, scopes: Iterable[tuple[str, date | None]]) -> None:
    unique_scopes = {(country_code, base_date) for country_code, base_date in scopes}
    for country_code, base_date in unique_scopes:
        _rebuild_aggregate_scope(db, country_code, base_date)


def _build_row_frame(df: pd.DataFrame) -> pd.DataFrame:
    level_series = (
        df["level"]
        if "level" in df.columns
        else pd.Series([None] * len(df), index=df.index, dtype="object")
    )
    warehouse_series = (
        df["warehouse"]
        if "warehouse" in df.columns
        else pd.Series([None] * len(df), index=df.index, dtype="object")
    )
    return pd.DataFrame(
        {
            "sku": df["sku"].fillna("").astype(str).str.strip(),
            "description": df["description"].fillna("").astype(str).str.strip(),
            "supplier": df["supplier"].fillna("").astype(str).str.strip(),
            "level": level_series.apply(_normalize_level_value),
            "warehouse": warehouse_series.map(_nullable_string),
            "quantity": pd.to_numeric(df["quantity"], errors="coerce").fillna(0).round().astype(int),
        }
    )


def _persist_dataframe(
    db: Session,
    *,
    df: pd.DataFrame,
    original_name: str,
    stored_name: str,
    s3_bucket: str,
    s3_key: str,
    country_code: str,
    base_date: date | None,
    size: int,
    client_id: str | None,
) -> tuple[dict, pd.DataFrame, tuple[str, date | None]]:
    uploaded_file = UploadedFile(
        original_name=original_name,
        stored_name=stored_name,
        s3_bucket=s3_bucket,
        s3_key=s3_key,
        country_type=_country_type(country_code),
        country_code=country_code,
        base_date=base_date,
        uploaded_by=None,
        size=size,
        status="parsed",
        error_message=None,
    )
    db.add(uploaded_file)
    db.flush()

    apply_db_locale_columns_to_raw_inventory_df(db, df, country_code)
    row_frame = _build_row_frame(df)
    for record in row_frame.to_dict(orient="records"):
        db.add(
            InventoryRow(
                uploaded_file_id=uploaded_file.id,
                sku=str(record["sku"]),
                description=_nullable_string(record["description"]),
                supplier=_nullable_string(record["supplier"]),
                level=_nullable_string(record["level"]),
                warehouse=_nullable_string(record["warehouse"]),
                quantity=int(record["quantity"]),
            )
        )

    aggregate_frame = row_frame.copy()
    aggregate_frame["country_code"] = country_code
    aggregate_frame["base_date"] = base_date
    return _serialize_uploaded_file(uploaded_file, client_id=client_id), aggregate_frame, (country_code, base_date)


def list_inventory_files(db: Session) -> list[dict]:
    uploaded_files = db.execute(
        select(UploadedFile).order_by(
            UploadedFile.country_code.asc(),
            UploadedFile.base_date.asc().nulls_last(),
            UploadedFile.uploaded_at.asc(),
        )
    ).scalars().all()
    return [_serialize_uploaded_file(uploaded_file) for uploaded_file in uploaded_files]


def patch_inventory_file_base_date(db: Session, file_id: str, base_date_value: str | None) -> dict:
    """데이터 관리에서 파일 기준일 변경 시 DB·집계(InventoryAggregate)를 즉시 맞춘다."""
    try:
        uid = uuid.UUID(str(file_id).strip())
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="file_id가 올바르지 않습니다.") from exc

    uploaded_file = db.get(UploadedFile, uid)
    if uploaded_file is None:
        raise HTTPException(status_code=404, detail="파일을 찾을 수 없습니다.")

    file_domain = getattr(uploaded_file, "file_domain", None) or "inventory"
    if file_domain != "inventory":
        raise HTTPException(status_code=400, detail="출고 파일은 기준일을 여기서 바꿀 수 없습니다.")

    old_scope = (uploaded_file.country_code, uploaded_file.base_date)

    if base_date_value is None or str(base_date_value).strip() == "":
        new_date = None
    else:
        raw = str(base_date_value).strip()[:10]
        try:
            new_date = date.fromisoformat(raw)
        except ValueError as exc:
            raise HTTPException(
                status_code=400,
                detail="날짜는 YYYY-MM-DD 형식이어야 합니다.",
            ) from exc

    if uploaded_file.base_date == new_date:
        return _serialize_uploaded_file(uploaded_file)

    uploaded_file.base_date = new_date
    try:
        db.flush()
        _rebuild_aggregate_scopes(db, [old_scope, (uploaded_file.country_code, new_date)])
        db.commit()
    except Exception:
        db.rollback()
        raise

    db.refresh(uploaded_file)
    return _serialize_uploaded_file(uploaded_file)


def get_inventory_view(db: Session, country_code: str) -> dict:
    normalized_country = str(country_code or "").strip().upper()
    if not normalized_country:
        raise HTTPException(status_code=400, detail="country_code가 필요합니다.")

    inv_domain = or_(
        InventoryAggregate.source_domain == "INVENTORY",
        InventoryAggregate.source_domain.is_(None),
    )
    rows = db.execute(
        select(
            InventoryAggregate.country_code,
            InventoryAggregate.base_date,
            InventoryAggregate.sku,
            InventoryAggregate.description,
            InventoryAggregate.supplier,
            InventoryAggregate.level,
            InventoryAggregate.warehouse,
            InventoryAggregate.total_quantity,
        ).where(
            InventoryAggregate.country_code == normalized_country,
            inv_domain,
        )
    ).all()

    if not rows:
        return {
            "summary": {"item_count": 0, "date_count": 0},
            "countries": [normalized_country],
            "dates": [],
            "rows": [],
        }

    frame = pd.DataFrame(
        [
            {
                "country": row.country_code,
                "date_key": row.base_date.isoformat() if row.base_date else "",
                "sku": row.sku,
                "description": row.description or "",
                "supplier": row.supplier or "",
                "level": row.level if row.level is not None else "__NONE__",
                "warehouse": row.warehouse or "",
                "quantity": int(row.total_quantity or 0),
            }
            for row in rows
            if row.base_date is not None
        ]
    )

    if frame.empty:
        return {
            "summary": {"item_count": 0, "date_count": 0},
            "countries": [normalized_country],
            "dates": [],
            "rows": [],
        }

    # 이미 저장된 집계라도 SKU·국가 기준 마스터 상품명·브랜드로 맞추고 동일 키는 수량 합산 (재업로드 없이 대시보드 정합)
    apply_db_locale_columns_to_inventory_merged(db, frame)
    frame = (
        frame.groupby(
            ["country", "date_key", "sku", "description", "supplier", "level", "warehouse"],
            as_index=False,
        )["quantity"]
        .sum()
    )

    pivot = (
        frame.pivot_table(
            index=["sku", "description", "supplier", "level", "warehouse", "country"],
            columns="date_key",
            values="quantity",
            aggfunc="sum",
            fill_value=0,
        )
        .reset_index()
        .rename_axis(None, axis=1)
    )
    pivot["level"] = pivot["level"].replace({"__NONE__": None})

    dates = sorted(frame["date_key"].dropna().unique().tolist())
    result_rows = pivot.to_dict(orient="records")
    attach_item_mapping_kr_to_inventory_rows(db, normalized_country, result_rows)
    return {
        "summary": {"item_count": len(result_rows), "date_count": len(dates)},
        "countries": [normalized_country],
        "dates": dates,
        "rows": result_rows,
    }


def persist_inventory_uploads(
    db: Session,
    files: list[UploadFile],
    file_dates: list[str] | None = None,
    file_countries: list[str] | None = None,
    file_client_ids: list[str] | None = None,
    file_db_ids: list[str] | None = None,
) -> list[dict]:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")
    if not is_s3_configured():
        raise HTTPException(status_code=500, detail="S3 환경변수가 아직 모두 설정되지 않았습니다.")
    if not settings.s3_bucket:
        raise HTTPException(status_code=500, detail="S3_BUCKET이 설정되지 않았습니다.")

    s3_client = get_s3_client()
    aggregation_version = _next_aggregation_version(db)
    affected_scopes: set[tuple[str, date | None]] = set()
    aggregate_frames: list[pd.DataFrame] = []
    persisted_files: list[dict] = []
    sku_allowlist_cache: dict[str, set[str]] = {}

    try:
        for idx, upload_file in enumerate(files):
            existing_file_id = ""
            if file_db_ids and idx < len(file_db_ids):
                existing_file_id = str(file_db_ids[idx] or "").strip()
            client_id = str(file_client_ids[idx] or "").strip() if file_client_ids and idx < len(file_client_ids) else None

            if existing_file_id:
                try:
                    parsed_existing_id = uuid.UUID(existing_file_id)
                except ValueError:
                    parsed_existing_id = None
                existing_file = db.get(UploadedFile, parsed_existing_id) if parsed_existing_id else None
                if existing_file:
                    persisted_files.append(_serialize_uploaded_file(existing_file, client_id=client_id))
                    continue

            raw_bytes = _read_upload_bytes(upload_file)
            if not raw_bytes:
                raise HTTPException(status_code=400, detail=f"{upload_file.filename}: 파일이 비어 있습니다.")

            df = _read_inventory_file(upload_file).copy()
            country_code = _resolve_country_code(df, upload_file, idx, file_countries)
            base_date = _resolve_base_date(df, upload_file, idx, file_dates, country_code)
            row_frame = _build_row_frame(df)
            validate_inventory_row_frame_item_mapping(
                db,
                row_frame,
                country_code,
                str(upload_file.filename or "파일"),
                allowlist_cache=sku_allowlist_cache,
            )

            stored_name = f"{uuid.uuid4()}{os.path.splitext(upload_file.filename or '')[1]}"
            s3_key = _build_s3_key(
                prefix=settings.s3_prefix,
                country_type=_country_type(country_code),
                country_code=country_code,
                base_date=base_date,
                stored_name=stored_name,
            )

            s3_client.put_object(
                Bucket=settings.s3_bucket,
                Key=s3_key,
                Body=raw_bytes,
                ContentType=upload_file.content_type or "application/octet-stream",
            )
            persisted_file, aggregate_frame, affected_scope = _persist_dataframe(
                db,
                df=df,
                original_name=upload_file.filename,
                stored_name=stored_name,
                s3_bucket=settings.s3_bucket,
                s3_key=s3_key,
                country_code=country_code,
                base_date=base_date,
                size=len(raw_bytes),
                client_id=client_id,
            )
            persisted_files.append(persisted_file)
            aggregate_frames.append(aggregate_frame)
            affected_scopes.add(affected_scope)

        if affected_scopes:
            scope_filters = []
            for country_code, base_date in affected_scopes:
                if base_date is None:
                    scope_filters.append(
                        and_(
                            InventoryAggregate.country_code == country_code,
                            InventoryAggregate.base_date.is_(None),
                        )
                    )
                else:
                    scope_filters.append(
                        and_(
                            InventoryAggregate.country_code == country_code,
                            InventoryAggregate.base_date == base_date,
                        )
                    )

            db.execute(delete(InventoryAggregate).where(or_(*scope_filters)))

        if aggregate_frames:
            merged = pd.concat(aggregate_frames, ignore_index=True)
            grouped = (
                merged.groupby(
                    [
                        "country_code",
                        "base_date",
                        "sku",
                        "description",
                        "supplier",
                        "level",
                        "warehouse",
                    ],
                    dropna=False,
                    as_index=False,
                )["quantity"]
                .sum()
            )

            for record in grouped.to_dict(orient="records"):
                db.add(
                    InventoryAggregate(
                        country_code=str(record["country_code"]),
                        base_date=record["base_date"],
                        sku=str(record["sku"]),
                        description=_nullable_string(record["description"]),
                        supplier=_nullable_string(record["supplier"]),
                        level=_nullable_string(record["level"]),
                        warehouse=_nullable_string(record["warehouse"]),
                        total_quantity=int(record["quantity"]),
                        aggregation_version=aggregation_version,
                        source_domain="INVENTORY",
                    )
                )

        db.commit()
        return persisted_files
    except Exception:
        db.rollback()
        raise


def prepare_inventory_direct_uploads(files: list[dict], db: Session | None = None) -> list[dict]:
    settings = get_runtime_settings()
    if not settings.s3_bucket or not is_s3_configured():
        raise HTTPException(status_code=500, detail="S3 환경변수가 아직 모두 설정되지 않았습니다.")

    prepared_files: list[dict] = []
    for file in files:
        client_id = str(file.get("client_id") or "").strip()
        original_name = str(file.get("original_name") or "").strip()
        country_code = _normalize_country_code(file.get("country_code")) or "KR"
        date_value = _parse_optional_date(file.get("date"), original_name)
        content_type = str(file.get("content_type") or "").strip() or "application/octet-stream"
        size = int(file.get("size") or 0)

        if db is not None:
            existing_file = db.execute(
                select(UploadedFile).where(
                    UploadedFile.original_name == original_name,
                    UploadedFile.country_code == country_code,
                    UploadedFile.base_date == date_value,
                )
            ).scalars().first()
            if existing_file is not None:
                prepared = _serialize_uploaded_file(existing_file, client_id=client_id)
                prepared_files.append(
                    {
                        **prepared,
                        "original_name": original_name,
                        "stored_name": existing_file.stored_name,
                        "s3_key": existing_file.s3_key,
                        "upload_url": "",
                        "size": existing_file.size or size,
                    }
                )
                continue

        stored_name = f"{uuid.uuid4()}{os.path.splitext(original_name)[1]}"
        s3_key = _build_s3_key(
            prefix=settings.s3_prefix,
            country_type=_country_type(country_code),
            country_code=country_code,
            base_date=date_value,
            stored_name=stored_name,
        )
        prepared_files.append(
            {
                "client_id": client_id,
                "original_name": original_name,
                "country_code": country_code,
                "date": date_value.isoformat() if date_value else "",
                "stored_name": stored_name,
                "s3_key": s3_key,
                "upload_url": create_presigned_upload_url(
                    bucket=settings.s3_bucket,
                    key=s3_key,
                    content_type=content_type,
                ),
                "file_id": None,
                "size": size,
            }
        )
    return prepared_files


def complete_inventory_direct_uploads(db: Session, files: list[dict]) -> list[dict]:
    settings = get_runtime_settings()
    if not settings.s3_bucket or not is_s3_configured():
        raise HTTPException(status_code=500, detail="S3 환경변수가 아직 모두 설정되지 않았습니다.")

    s3_client = get_s3_client()
    aggregation_version = _next_aggregation_version(db)
    affected_scopes: set[tuple[str, date | None]] = set()
    aggregate_frames: list[pd.DataFrame] = []
    persisted_files: list[dict] = []
    sku_allowlist_cache: dict[str, set[str]] = {}

    try:
        for file in files:
            file_id = str(file.get("file_id") or "").strip()
            client_id = str(file.get("client_id") or "").strip() or None
            if file_id:
                existing_file = db.get(UploadedFile, uuid.UUID(file_id))
                if existing_file is not None:
                    persisted_files.append(_serialize_uploaded_file(existing_file, client_id=client_id))
                    continue

            original_name = str(file.get("original_name") or "").strip()
            if not original_name:
                raise HTTPException(status_code=400, detail="original_name이 필요합니다.")
            country_code = _normalize_country_code(file.get("country_code")) or "KR"
            base_date_override = _parse_optional_date(file.get("date"), original_name)
            stored_name = str(file.get("stored_name") or "").strip()
            s3_key = str(file.get("s3_key") or "").strip()
            size = int(file.get("size") or 0)
            if not stored_name or not s3_key:
                raise HTTPException(status_code=400, detail=f"{original_name}: 업로드 완료 정보가 부족합니다.")

            existing_file = db.execute(
                select(UploadedFile).where(
                    UploadedFile.original_name == original_name,
                    UploadedFile.country_code == country_code,
                    UploadedFile.base_date == base_date_override,
                    UploadedFile.s3_key == s3_key,
                )
            ).scalars().first()
            if existing_file is not None:
                persisted_files.append(_serialize_uploaded_file(existing_file, client_id=client_id))
                continue

            response = s3_client.get_object(Bucket=settings.s3_bucket, Key=s3_key)
            raw_bytes = response["Body"].read()
            df = _read_inventory_bytes(original_name, raw_bytes).copy()
            row_frame = _build_row_frame(df)
            validate_inventory_row_frame_item_mapping(
                db,
                row_frame,
                country_code,
                original_name,
                allowlist_cache=sku_allowlist_cache,
            )
            base_date = base_date_override or _resolve_base_date(df, type("Tmp", (), {"filename": original_name})(), 0, None, country_code)
            persisted_file, aggregate_frame, affected_scope = _persist_dataframe(
                db,
                df=df,
                original_name=original_name,
                stored_name=stored_name,
                s3_bucket=settings.s3_bucket,
                s3_key=s3_key,
                country_code=country_code,
                base_date=base_date,
                size=size or len(raw_bytes),
                client_id=client_id,
            )
            persisted_files.append(persisted_file)
            aggregate_frames.append(aggregate_frame)
            affected_scopes.add(affected_scope)

        if affected_scopes:
            scope_filters = []
            for country_code, base_date in affected_scopes:
                scope_filters.append(_scope_filter(country_code, base_date))
            db.execute(delete(InventoryAggregate).where(or_(*scope_filters)))

        if aggregate_frames:
            merged = pd.concat(aggregate_frames, ignore_index=True)
            grouped = (
                merged.groupby(
                    ["country_code", "base_date", "sku", "description", "supplier", "level", "warehouse"],
                    dropna=False,
                    as_index=False,
                )["quantity"]
                .sum()
            )
            for record in grouped.to_dict(orient="records"):
                db.add(
                    InventoryAggregate(
                        country_code=str(record["country_code"]),
                        base_date=record["base_date"],
                        sku=str(record["sku"]),
                        description=_nullable_string(record["description"]),
                        supplier=_nullable_string(record["supplier"]),
                        level=_nullable_string(record["level"]),
                        warehouse=_nullable_string(record["warehouse"]),
                        total_quantity=int(record["quantity"]),
                        aggregation_version=aggregation_version,
                        source_domain="INVENTORY",
                    )
                )

        db.commit()
        return persisted_files
    except Exception:
        db.rollback()
        raise


def delete_inventory_file(db: Session, file_id: str) -> None:
    try:
        parsed_id = uuid.UUID(str(file_id))
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="올바른 파일 ID 형식이 아닙니다.") from exc

    uploaded_file = db.get(UploadedFile, parsed_id)
    if uploaded_file is None:
        raise HTTPException(status_code=404, detail="삭제할 파일을 찾을 수 없습니다.")

    b, k = _resolved_upload_s3_bucket_and_key(uploaded_file)
    _delete_s3_object(b, k)

    affected_scope = (uploaded_file.country_code, uploaded_file.base_date)
    file_domain = getattr(uploaded_file, "file_domain", None) or "inventory"
    country_u = str(uploaded_file.country_code or "").strip().upper()
    # 레거시: country_code만 SHIPMENT 이고 file_domain 이 기본 inventory 인 행이 있음 → 매트릭스 purge 누락 방지
    is_shipment_cleanup = file_domain == "shipment" or country_u == "SHIPMENT"
    db.delete(uploaded_file)
    db.flush()
    if is_shipment_cleanup:
        from domains.inventory.services.shipment_persistence_service import (
            _rebuild_all_shipment_aggregates,
            purge_shipment_matrix_rows_if_no_matrix_uploads,
        )

        purge_shipment_matrix_rows_if_no_matrix_uploads(db)
        _rebuild_all_shipment_aggregates(db)
    else:
        _rebuild_aggregate_scopes(db, [affected_scope])
    db.commit()


def delete_inventory_files(db: Session, country_code: str | None = None) -> dict[str, int]:
    settings = get_runtime_settings()
    raw_cc = str(country_code or "").strip()
    scope_country = _normalize_country_code(raw_cc) if raw_cc else None
    if raw_cc and scope_country is None:
        raise HTTPException(status_code=400, detail="올바른 country_code가 아닙니다.")

    query = select(UploadedFile)
    if scope_country:
        query = query.where(UploadedFile.country_code == scope_country)

    uploaded_files = list(db.execute(query).scalars().all())
    if not uploaded_files:
        # DB 행이 없어도 S3에 고아 객체가 있을 수 있음 → 스윕만 수행
        if is_s3_configured() and settings.s3_bucket:
            if scope_country:
                p = _s3_prefix_for_country_scope(settings, scope_country)
                if p:
                    _delete_s3_prefix(settings.s3_bucket, p)
            else:
                p = _s3_prefix_inventory_root(settings)
                if p:
                    _delete_s3_prefix(settings.s3_bucket, p)
        return {"deleted_count": 0}

    ship_touched = any(
        (getattr(f, "file_domain", None) or "inventory") == "shipment"
        or str(f.country_code or "").strip().upper() == "SHIPMENT"
        for f in uploaded_files
    )
    inv_scopes_unique = {
        (f.country_code, f.base_date)
        for f in uploaded_files
        if (getattr(f, "file_domain", None) or "inventory") != "shipment"
        and str(f.country_code or "").strip().upper() != "SHIPMENT"
    }
    for uploaded_file in uploaded_files:
        b, k = _resolved_upload_s3_bucket_and_key(uploaded_file)
        _delete_s3_object(b, k)
        db.delete(uploaded_file)

    db.flush()
    if inv_scopes_unique:
        _rebuild_aggregate_scopes(db, list(inv_scopes_unique))
    if ship_touched:
        from domains.inventory.services.shipment_persistence_service import (
            _rebuild_all_shipment_aggregates,
            purge_shipment_matrix_rows_if_no_matrix_uploads,
        )

        purge_shipment_matrix_rows_if_no_matrix_uploads(db)
        _rebuild_all_shipment_aggregates(db)
    db.commit()

    if is_s3_configured() and settings.s3_bucket:
        if scope_country:
            sweep = _s3_prefix_for_country_scope(settings, scope_country)
            if sweep:
                _delete_s3_prefix(settings.s3_bucket, sweep)
        else:
            sweep = _s3_prefix_inventory_root(settings)
            if sweep:
                _delete_s3_prefix(settings.s3_bucket, sweep)

    return {"deleted_count": len(uploaded_files)}
