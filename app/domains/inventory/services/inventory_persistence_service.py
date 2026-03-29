import os
import re
import uuid
from collections.abc import Iterable
from datetime import date

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import and_, delete, func, or_, select
from sqlalchemy.orm import Session

from app.domains.inventory.models import InventoryAggregate, InventoryRow, UploadedFile
from app.domains.inventory.services.inventory_aggregate_service import (
    _detect_country,
    _normalize_country_code,
    _normalize_item_code,
    _normalize_level_value,
    _parse_date_from_filename,
    _parse_date_series,
    _read_inventory_file,
)
from app.shared.config import get_runtime_settings
from app.shared.storage import get_s3_client, is_s3_configured


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


def _extract_service_item_code(country_code: str, raw_item_code: str) -> str:
    normalized = _normalize_item_code(raw_item_code)
    if re.fullmatch(r"\d{5}", normalized):
        return normalized

    prefixed_match = re.search(r"^[A-Z]+(\d{5})", normalized)
    if prefixed_match:
        return prefixed_match.group(1)

    if country_code == "KR":
        return normalized

    match = re.search(r"(\d{5})", normalized)
    if match:
        return match.group(1)

    digits = re.sub(r"\D", "", normalized)
    if len(digits) >= 5:
        return digits[:5]
    return normalized


def _build_s3_key(prefix: str, country_type: str, country_code: str, base_date: date | None, stored_name: str) -> str:
    parts = [
        prefix.strip("/"),
        country_type.lower(),
        country_code.lower(),
        base_date.isoformat() if base_date else "unknown-date",
        stored_name,
    ]
    return "/".join(part for part in parts if part)


def _nullable_string(value: object) -> str | None:
    if pd.isna(value):
        return None

    raw = str(value).strip()
    return raw or None


def _next_aggregation_version(db: Session) -> int:
    current = db.execute(select(func.max(InventoryAggregate.aggregation_version))).scalar_one_or_none()
    return int(current or 0) + 1


def _serialize_uploaded_file(uploaded_file: UploadedFile, client_id: str | None = None) -> dict:
    return {
        "client_id": client_id,
        "file_id": str(uploaded_file.id),
        "name": uploaded_file.original_name,
        "country": uploaded_file.country_code,
        "date": uploaded_file.base_date.isoformat() if uploaded_file.base_date else "",
        "size": uploaded_file.size or 0,
        "status": uploaded_file.status,
    }


def _scope_filter(country_code: str, base_date: date | None):
    if base_date is None:
        return and_(InventoryAggregate.country_code == country_code, InventoryAggregate.base_date.is_(None))
    return and_(InventoryAggregate.country_code == country_code, InventoryAggregate.base_date == base_date)


def _uploaded_file_scope_filter(country_code: str, base_date: date | None):
    if base_date is None:
        return and_(UploadedFile.country_code == country_code, UploadedFile.base_date.is_(None))
    return and_(UploadedFile.country_code == country_code, UploadedFile.base_date == base_date)


def _delete_s3_object(bucket: str | None, key: str | None) -> None:
    if not bucket or not key or not is_s3_configured():
        return
    client = get_s3_client()
    client.delete_object(Bucket=bucket, Key=key)


def _rebuild_aggregate_scope(db: Session, country_code: str, base_date: date | None) -> None:
    db.execute(delete(InventoryAggregate).where(_scope_filter(country_code, base_date)))

    uploaded_file_ids = db.execute(
        select(UploadedFile.id).where(_uploaded_file_scope_filter(country_code, base_date))
    ).scalars().all()
    if not uploaded_file_ids:
        return

    rows = db.execute(
        select(
            InventoryRow.raw_item_code,
            InventoryRow.item_code,
            InventoryRow.description,
            InventoryRow.supplier,
            InventoryRow.level,
            InventoryRow.quantity,
        ).where(InventoryRow.uploaded_file_id.in_(uploaded_file_ids))
    ).all()
    if not rows:
        return

    frame = pd.DataFrame(
        [
            {
                "raw_item_code": row.raw_item_code,
                "item_code": row.item_code,
                "description": row.description,
                "supplier": row.supplier,
                "level": row.level,
                "quantity": int(row.quantity or 0),
            }
            for row in rows
        ]
    )
    grouped = (
        frame.groupby(
            ["raw_item_code", "item_code", "description", "supplier", "level"],
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
                raw_item_code=str(record["raw_item_code"]),
                item_code=str(record["item_code"]),
                description=_nullable_string(record["description"]),
                supplier=_nullable_string(record["supplier"]),
                level=_nullable_string(record["level"]),
                total_quantity=int(record["quantity"]),
                aggregation_version=aggregation_version,
            )
        )


def _rebuild_aggregate_scopes(db: Session, scopes: Iterable[tuple[str, date | None]]) -> None:
    unique_scopes = {(country_code, base_date) for country_code, base_date in scopes}
    for country_code, base_date in unique_scopes:
        _rebuild_aggregate_scope(db, country_code, base_date)


def list_inventory_files(db: Session) -> list[dict]:
    uploaded_files = db.execute(
        select(UploadedFile).order_by(
            UploadedFile.country_code.asc(),
            UploadedFile.base_date.asc().nulls_last(),
            UploadedFile.uploaded_at.asc(),
        )
    ).scalars().all()
    return [_serialize_uploaded_file(uploaded_file) for uploaded_file in uploaded_files]


def get_inventory_view(db: Session, country_code: str) -> dict:
    normalized_country = str(country_code or "").strip().upper()
    if not normalized_country:
        raise HTTPException(status_code=400, detail="country_code가 필요합니다.")

    rows = db.execute(
        select(
            InventoryAggregate.country_code,
            InventoryAggregate.base_date,
            InventoryAggregate.raw_item_code,
            InventoryAggregate.description,
            InventoryAggregate.supplier,
            InventoryAggregate.level,
            InventoryAggregate.total_quantity,
        ).where(InventoryAggregate.country_code == normalized_country)
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
                "itemno": row.raw_item_code,
                "description": row.description or "",
                "supplier": row.supplier or "",
                "category": "",
                "level": row.level if row.level is not None else "__NONE__",
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

    pivot = (
        frame.pivot_table(
            index=["itemno", "description", "supplier", "category", "level", "country"],
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
            country_type = _country_type(country_code)

            stored_name = f"{uuid.uuid4()}{os.path.splitext(upload_file.filename or '')[1]}"
            s3_key = _build_s3_key(
                prefix=settings.s3_prefix,
                country_type=country_type,
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

            uploaded_file = UploadedFile(
                original_name=upload_file.filename,
                stored_name=stored_name,
                s3_bucket=settings.s3_bucket,
                s3_key=s3_key,
                country_type=country_type,
                country_code=country_code,
                base_date=base_date,
                uploaded_by=None,
                size=len(raw_bytes),
                status="parsed",
                error_message=None,
            )
            db.add(uploaded_file)
            db.flush()
            persisted_files.append(_serialize_uploaded_file(uploaded_file, client_id=client_id))

            level_series = (
                df["level"]
                if "level" in df.columns
                else pd.Series([None] * len(df), index=df.index, dtype="object")
            )

            row_frame = pd.DataFrame(
                {
                    "raw_item_code": df["raw_item_code"].fillna("").astype(str).str.strip(),
                    "item_code": df["raw_item_code"]
                    .fillna("")
                    .astype(str)
                    .map(lambda value: _extract_service_item_code(country_code, value)),
                    "description": df["description"].fillna("").astype(str).str.strip(),
                    "supplier": df["supplier"].fillna("").astype(str).str.strip(),
                    "level": level_series.apply(_normalize_level_value),
                    "quantity": pd.to_numeric(df["quantity"], errors="coerce").fillna(0).round().astype(int),
                }
            )

            for record in row_frame.to_dict(orient="records"):
                db.add(
                    InventoryRow(
                        uploaded_file_id=uploaded_file.id,
                        raw_item_code=str(record["raw_item_code"]),
                        item_code=str(record["item_code"]),
                        description=_nullable_string(record["description"]),
                        supplier=_nullable_string(record["supplier"]),
                        level=_nullable_string(record["level"]),
                        quantity=int(record["quantity"]),
                    )
                )

            aggregate_frame = row_frame.copy()
            aggregate_frame["country_code"] = country_code
            aggregate_frame["base_date"] = base_date
            aggregate_frames.append(aggregate_frame)
            affected_scopes.add((country_code, base_date))

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
                        "raw_item_code",
                        "item_code",
                        "description",
                        "supplier",
                        "level",
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
                        raw_item_code=str(record["raw_item_code"]),
                        item_code=str(record["item_code"]),
                        description=_nullable_string(record["description"]),
                        supplier=_nullable_string(record["supplier"]),
                        level=_nullable_string(record["level"]),
                        total_quantity=int(record["quantity"]),
                        aggregation_version=aggregation_version,
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

    _delete_s3_object(uploaded_file.s3_bucket, uploaded_file.s3_key)

    affected_scope = (uploaded_file.country_code, uploaded_file.base_date)
    db.delete(uploaded_file)
    db.flush()
    _rebuild_aggregate_scopes(db, [affected_scope])
    db.commit()


def delete_inventory_files(db: Session, country_code: str | None = None) -> dict[str, int]:
    query = select(UploadedFile)
    if country_code:
        query = query.where(UploadedFile.country_code == country_code)

    uploaded_files = list(db.execute(query).scalars().all())
    if not uploaded_files:
        return {"deleted_count": 0}

    scopes = [(file.country_code, file.base_date) for file in uploaded_files]
    for uploaded_file in uploaded_files:
        _delete_s3_object(uploaded_file.s3_bucket, uploaded_file.s3_key)
        db.delete(uploaded_file)

    db.flush()
    _rebuild_aggregate_scopes(db, scopes)
    db.commit()
    return {"deleted_count": len(uploaded_files)}
