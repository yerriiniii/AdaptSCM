"""출고 파일 S3·DB 저장 및 집계 조회."""

from __future__ import annotations

import os
import uuid
from datetime import date, datetime, timezone
from typing import Any

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import delete, func, literal, select
from sqlalchemy.orm import Session

from domains.inventory.models import InventoryAggregate, InventoryRow, ShipmentMatrixRow, UploadedFile
from domains.inventory.services.inventory_persistence_service import (
    _next_aggregation_version,
    _nullable_string,
    _serialize_uploaded_file,
)
from domains.inventory.services.shipment_aggregate_service import (
    SHIPMENT_CHANNEL_LABELS,
    build_shipment_wide_dashboard,
)
from domains.inventory.services.shipment_matrix_service import (
    DEFAULT_SHIPMENT_MATRIX_YEAR,
    build_shipment_matrix_view,
    parse_shipment_matrix_workbook,
)
from shared.config import get_runtime_settings
from shared.storage import get_s3_client, is_s3_configured

SOURCE_SHIPMENT = "SHIPMENT"
FILE_DOMAIN_SHIPMENT = "shipment"
SHIPMENT_BUCKET_COUNTRY = "SHIPMENT"
SHIPMENT_MATRIX_UPLOAD_KEY = "matrix"


def _normalize_shipment_order_at(val: Any) -> datetime:
    """엑셀 주문일·DB 값 → 중복 비교용(초 단위, 마이크로초·tz 제거)."""
    if isinstance(val, datetime):
        dt = val
    elif isinstance(val, date) and not isinstance(val, datetime):
        dt = datetime.combine(val, datetime.min.time())
    else:
        dt = pd.to_datetime(val).to_pydatetime()
    if dt.tzinfo is not None:
        dt = dt.replace(tzinfo=None)
    return dt.replace(microsecond=0)


def _load_existing_shipment_order_times_for_sheet(
    db: Session, sheet_key: str, order_times: list[datetime]
) -> set[datetime]:
    """같은 시트 키로 이미 저장된 주문일 시각(초 단위)."""
    clean = [_normalize_shipment_order_at(t) for t in order_times if t is not None]
    if not sheet_key or not clean:
        return set()
    t_min, t_max = min(clean), max(clean)
    stmt = (
        select(InventoryRow.shipment_order_at)
        .join(UploadedFile, InventoryRow.uploaded_file_id == UploadedFile.id)
        .where(
            UploadedFile.file_domain == FILE_DOMAIN_SHIPMENT,
            UploadedFile.shipment_sheet_key == sheet_key,
            InventoryRow.shipment_order_at.isnot(None),
            InventoryRow.shipment_order_at >= t_min,
            InventoryRow.shipment_order_at <= t_max,
        )
    )
    return {_normalize_shipment_order_at(t) for t in db.execute(stmt).scalars().all() if t is not None}


def _shipment_s3_key(prefix: str, stored_name: str) -> str:
    root = (prefix or "inventory").strip().strip("/")
    return f"{root}/shipment/{stored_name}"


def purge_shipment_matrix_rows_if_no_matrix_uploads(db: Session) -> None:
    """`shipment_sheet_key == matrix` 인 업로드가 하나도 없으면 `shipment_matrix_rows` 전부 삭제.

    - 매트릭스 엑셀만 이 키를 씀. 롱포맷 출고는 다른 키·`inventory_rows`만 사용.
    - 예전 데이터는 `file_domain`이 `inventory`로 남아 삭제 시 출고 분기를 타지 않는 경우가 있어,
      호출부에서 `country_code == SHIPMENT` 삭제도 출고 정리로 묶음.
    - `file_domain` 조건은 넣지 않음(도메인 누락 레거시와 동일하게 matrix 키만 보면 됨).
    """
    n = db.scalar(
        select(func.count()).select_from(UploadedFile).where(
            UploadedFile.shipment_sheet_key == SHIPMENT_MATRIX_UPLOAD_KEY,
        )
    )
    if (n or 0) == 0:
        db.execute(delete(ShipmentMatrixRow))


def _rebuild_all_shipment_aggregates(db: Session) -> None:
    """출고 inventory_rows 전량을 메모리에 올리지 않고 DB에서 GROUP BY 후 집계 행만 적재."""
    db.execute(delete(InventoryAggregate).where(InventoryAggregate.source_domain == SOURCE_SHIPMENT))

    co_country = func.coalesce(InventoryRow.row_country_code, literal("KR"))
    stmt = (
        select(
            co_country.label("country_code"),
            InventoryRow.row_snapshot_date.label("base_date"),
            InventoryRow.sku,
            InventoryRow.description,
            InventoryRow.supplier,
            InventoryRow.level,
            InventoryRow.warehouse,
            func.sum(InventoryRow.quantity).label("total_quantity"),
        )
        .join(UploadedFile, InventoryRow.uploaded_file_id == UploadedFile.id)
        .where(
            UploadedFile.file_domain == FILE_DOMAIN_SHIPMENT,
            InventoryRow.row_snapshot_date.isnot(None),
        )
        .group_by(
            co_country,
            InventoryRow.row_snapshot_date,
            InventoryRow.sku,
            InventoryRow.description,
            InventoryRow.supplier,
            InventoryRow.level,
            InventoryRow.warehouse,
        )
    )
    rows = db.execute(stmt).all()
    if not rows:
        return

    ver = _next_aggregation_version(db)
    now = datetime.now(timezone.utc)
    mappings: list[dict[str, Any]] = []
    for r in rows:
        bd = r.base_date
        if bd is None:
            continue
        if hasattr(bd, "date"):
            bd = bd.date()
        qty = r.total_quantity
        try:
            qty_i = int(round(float(qty)))
        except (TypeError, ValueError):
            qty_i = 0
        mappings.append(
            {
                "id": uuid.uuid4(),
                "country_code": str(r.country_code or "KR"),
                "base_date": bd,
                "sku": str(r.sku or ""),
                "description": _nullable_string(r.description),
                "supplier": _nullable_string(r.supplier),
                "level": _nullable_string(r.level),
                "warehouse": _nullable_string(r.warehouse),
                "total_quantity": qty_i,
                "aggregation_version": ver,
                "uploaded_at": now,
                "source_domain": SOURCE_SHIPMENT,
            }
        )
    if mappings:
        db.bulk_insert_mappings(InventoryAggregate, mappings)


def persist_shipment_uploads(
    db: Session,
    file_payloads: list[tuple[UploadFile, bytes, str, list[dict[str, Any]]]],
) -> list[dict]:
    """파싱된 롱 레코드와 원본 바이트로 S3·inventory_rows 저장(엑셀 재파싱 없음).

    동일 「N월 출고 ALL」시트에서 **주문일(일시·초)** 이 DB에 이미 있으면 스킵하고, 없는 시각만 적재한다.
    같은 통합 요청 안에서도 동일 주문일이 두 번 나오면 한 번만 넣는다.
    """
    settings = get_runtime_settings()
    if not settings.database_enabled:
        return []
    if not is_s3_configured() or not settings.s3_bucket:
        raise HTTPException(status_code=500, detail="S3가 설정되지 않아 출고 파일을 저장할 수 없습니다.")

    s3 = get_s3_client()
    persisted: list[dict] = []
    row_ts = datetime.now(timezone.utc)
    try:
        for upload_file, raw, sheet_key, records in file_payloads:
            if not records:
                continue
            batch_times = [rec["order_at"] for rec in records if rec.get("order_at") is not None]
            seen = _load_existing_shipment_order_times_for_sheet(db, sheet_key, batch_times)
            new_records: list[dict[str, Any]] = []
            for rec in records:
                oa = _normalize_shipment_order_at(rec["order_at"])
                if oa in seen:
                    continue
                seen.add(oa)
                new_records.append(rec)
            if not new_records:
                continue
            records = new_records

            stored_name = f"{uuid.uuid4()}{os.path.splitext(upload_file.filename or '')[1] or '.xlsx'}"
            s3_key = _shipment_s3_key(settings.s3_prefix or "inventory", stored_name)
            s3.put_object(
                Bucket=settings.s3_bucket,
                Key=s3_key,
                Body=raw,
                ContentType=upload_file.content_type or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            uploaded_file = UploadedFile(
                original_name=upload_file.filename or stored_name,
                stored_name=stored_name,
                s3_bucket=settings.s3_bucket,
                s3_key=s3_key,
                country_type="SHIPMENT",
                country_code=SHIPMENT_BUCKET_COUNTRY,
                base_date=None,
                uploaded_by=None,
                size=len(raw),
                status="parsed",
                error_message=None,
                file_domain=FILE_DOMAIN_SHIPMENT,
                shipment_sheet_key=sheet_key or None,
            )
            db.add(uploaded_file)
            db.flush()

            row_maps: list[dict[str, Any]] = []
            for rec in records:
                d = rec["date"]
                if not isinstance(d, date):
                    d = pd.to_datetime(d).date()
                row_maps.append(
                    {
                        "id": uuid.uuid4(),
                        "uploaded_file_id": uploaded_file.id,
                        "sku": str(rec.get("sku") or "").strip(),
                        "description": _nullable_string(rec.get("description")),
                        "supplier": _nullable_string(rec.get("brand")),
                        "level": _nullable_string(rec.get("level")),
                        "warehouse": _nullable_string(rec.get("channel")),
                        "vendor_raw": _nullable_string(rec.get("vendor_raw")),
                        "shipment_order_at": rec["order_at"],
                        "quantity": float(int(rec.get("quantity") or 0)),
                        "uploaded_at": row_ts,
                        "row_snapshot_date": d,
                        "row_country_code": str(rec.get("country") or "KR").strip().upper() or "KR",
                    }
                )
            if row_maps:
                db.bulk_insert_mappings(InventoryRow, row_maps)
            persisted.append(_serialize_uploaded_file(uploaded_file))

        _rebuild_all_shipment_aggregates(db)
        db.commit()
        return persisted
    except Exception:
        db.rollback()
        raise


def _upsert_shipment_matrix_channel_rows(
    db: Session,
    channel_sheet: str,
    data_year: int,
    recs: list[dict[str, Any]],
) -> None:
    for rec in recs:
        sku = str(rec.get("sku") or "").strip()
        if not sku:
            continue
        mt_new = {str(k): int(v) for k, v in (rec.get("monthly") or {}).items()}
        dt_new = {str(k): int(v) for k, v in (rec.get("daily") or {}).items()}
        mkt_raw = rec.get("mkt_priority")
        cat_raw = rec.get("category")
        mkt = str(mkt_raw).strip() if mkt_raw is not None and str(mkt_raw).strip() else None
        cat = str(cat_raw).strip() if cat_raw is not None and str(cat_raw).strip() else None

        stmt = select(ShipmentMatrixRow).where(
            ShipmentMatrixRow.channel_sheet == channel_sheet,
            ShipmentMatrixRow.data_year == data_year,
            ShipmentMatrixRow.sku == sku,
        )
        existing = db.execute(stmt).scalar_one_or_none()
        if existing:
            mt = {str(k): int(v) for k, v in (existing.monthly_totals or {}).items()}
            for k, v in mt_new.items():
                mt[k] = v
            dt = {str(k): int(v) for k, v in (existing.daily_totals or {}).items()}
            for k, v in dt_new.items():
                dt[k] = v
            existing.monthly_totals = mt
            existing.daily_totals = dt
            existing.mkt_priority = mkt
            existing.category = cat
        else:
            db.add(
                ShipmentMatrixRow(
                    channel_sheet=channel_sheet,
                    data_year=data_year,
                    sku=sku,
                    mkt_priority=mkt,
                    category=cat,
                    monthly_totals=mt_new or None,
                    daily_totals=dt_new or None,
                )
            )


def persist_shipment_matrix_uploads(
    db: Session,
    files: list[tuple[UploadFile, bytes]],
    *,
    kr_sku_brand_name: dict[str, tuple[str, str]] | None = None,
    data_year: int = DEFAULT_SHIPMENT_MATRIX_YEAR,
) -> list[dict]:
    """매트릭스 출고 엑셀: S3 저장 후 채널(시트)별 shipment_matrix_rows 업서트."""
    settings = get_runtime_settings()
    if not settings.database_enabled:
        return []
    if not is_s3_configured() or not settings.s3_bucket:
        raise HTTPException(status_code=500, detail="S3가 설정되지 않아 출고 파일을 저장할 수 없습니다.")

    s3 = get_s3_client()
    persisted: list[dict] = []
    try:
        for upload_file, raw in files:
            parsed = parse_shipment_matrix_workbook(
                db,
                raw,
                upload_file.filename or "",
                kr_sku_brand_name=kr_sku_brand_name,
                data_year=data_year,
            )
            for channel_sheet, recs in parsed.items():
                if recs:
                    _upsert_shipment_matrix_channel_rows(db, channel_sheet, data_year, recs)

            stored_name = f"{uuid.uuid4()}{os.path.splitext(upload_file.filename or '')[1] or '.xlsx'}"
            s3_key = _shipment_s3_key(settings.s3_prefix or "inventory", stored_name)
            s3.put_object(
                Bucket=settings.s3_bucket,
                Key=s3_key,
                Body=raw,
                ContentType=upload_file.content_type or "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            uploaded_file = UploadedFile(
                original_name=upload_file.filename or stored_name,
                stored_name=stored_name,
                s3_bucket=settings.s3_bucket,
                s3_key=s3_key,
                country_type="SHIPMENT",
                country_code=SHIPMENT_BUCKET_COUNTRY,
                base_date=None,
                uploaded_by=None,
                size=len(raw),
                status="parsed",
                error_message=None,
                file_domain=FILE_DOMAIN_SHIPMENT,
                shipment_sheet_key=SHIPMENT_MATRIX_UPLOAD_KEY,
            )
            db.add(uploaded_file)
            db.flush()
            persisted.append(_serialize_uploaded_file(uploaded_file))

        db.commit()
        return persisted
    except Exception:
        db.rollback()
        raise


def get_shipment_view(
    db: Session,
    *,
    channel: str | None = None,
    data_year: int = DEFAULT_SHIPMENT_MATRIX_YEAR,
) -> dict:
    """출고 현황 매트릭스(칩=시트명). 첫 행에 합계 행 포함."""
    payload = build_shipment_matrix_view(db, channel=channel, data_year=data_year)
    totals_row = payload.pop("shipment_totals_row", None)
    body_rows: list[dict[str, Any]] = list(payload.get("rows") or [])
    if totals_row:
        payload["rows"] = [totals_row, *body_rows]
    else:
        payload["rows"] = body_rows
    return payload
