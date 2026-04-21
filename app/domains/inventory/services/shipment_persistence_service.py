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

from app.domains.inventory.models import InventoryAggregate, InventoryRow, UploadedFile
from app.domains.inventory.services.inventory_persistence_service import (
    _next_aggregation_version,
    _nullable_string,
    _serialize_uploaded_file,
)
from app.domains.inventory.services.shipment_aggregate_service import (
    SHIPMENT_CHANNEL_LABELS,
    build_shipment_wide_dashboard,
)
from app.shared.config import get_runtime_settings
from app.shared.storage import get_s3_client, is_s3_configured

SOURCE_SHIPMENT = "SHIPMENT"
FILE_DOMAIN_SHIPMENT = "shipment"
SHIPMENT_BUCKET_COUNTRY = "SHIPMENT"


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


def get_shipment_view(db: Session) -> dict:
    """원본 출고 행(판매처 원문 포함)으로 복원 — 집계 테이블에는 판매처가 없음.

    DB에서 행 단위 전량을 앱으로 올리면 대용량 시 타임아웃되므로, 동일 키는 SQL에서 SUM으로 묶어
    전송량·파이썬 루프 비용을 줄인다(와이드 대시보드 결과는 동일하게 합산됨).
    """
    rows = db.execute(
        select(
            InventoryRow.row_snapshot_date,
            InventoryRow.sku,
            InventoryRow.description,
            InventoryRow.supplier,
            InventoryRow.level,
            InventoryRow.warehouse,
            func.sum(InventoryRow.quantity).label("quantity_sum"),
            InventoryRow.vendor_raw,
            InventoryRow.row_country_code,
        )
        .join(UploadedFile, InventoryRow.uploaded_file_id == UploadedFile.id)
        .where(
            UploadedFile.file_domain == FILE_DOMAIN_SHIPMENT,
            InventoryRow.row_snapshot_date.isnot(None),
        )
        .group_by(
            InventoryRow.row_snapshot_date,
            InventoryRow.sku,
            InventoryRow.description,
            InventoryRow.supplier,
            InventoryRow.level,
            InventoryRow.warehouse,
            InventoryRow.vendor_raw,
            InventoryRow.row_country_code,
        )
    ).all()

    if not rows:
        return {
            "summary": {"item_count": 0, "date_count": 0},
            "countries": [],
            "dates": [],
            "channels": list(SHIPMENT_CHANNEL_LABELS),
            "rows": [],
        }

    long_recs: list[dict[str, Any]] = []
    for row in rows:
        if row.row_snapshot_date is None:
            continue
        try:
            qty_i = int(round(float(getattr(row, "quantity_sum", None) or 0)))
        except (TypeError, ValueError):
            qty_i = 0
        long_recs.append(
            {
                "country": str(row.row_country_code or "KR").strip().upper() or "KR",
                "channel": row.warehouse or "",
                "sku": row.sku or "",
                "description": row.description or "",
                "brand": row.supplier or "",
                "date": row.row_snapshot_date,
                "quantity": qty_i,
                "level": row.level,
                "vendor_raw": (row.vendor_raw or "").strip() if row.vendor_raw is not None else "",
            }
        )
    if not long_recs:
        return {
            "summary": {"item_count": 0, "date_count": 0},
            "countries": [],
            "dates": [],
            "channels": list(SHIPMENT_CHANNEL_LABELS),
            "rows": [],
        }

    return build_shipment_wide_dashboard(long_recs)
