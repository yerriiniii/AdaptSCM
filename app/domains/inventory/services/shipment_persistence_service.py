"""출고 파일 S3·DB 저장 및 집계 조회."""

from __future__ import annotations

import os
import uuid
from datetime import date, datetime, timezone
from typing import Any

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import delete, select
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


def _shipment_s3_key(prefix: str, stored_name: str) -> str:
    root = (prefix or "inventory").strip().strip("/")
    return f"{root}/shipment/{stored_name}"


def _rebuild_all_shipment_aggregates(db: Session) -> None:
    db.execute(delete(InventoryAggregate).where(InventoryAggregate.source_domain == SOURCE_SHIPMENT))

    # ORM 전체 로드 대신 필요한 컬럼만 조회 + 집계 행은 bulk insert
    row_tuples = db.execute(
        select(
            InventoryRow.sku,
            InventoryRow.description,
            InventoryRow.supplier,
            InventoryRow.level,
            InventoryRow.warehouse,
            InventoryRow.quantity,
            InventoryRow.row_snapshot_date,
            InventoryRow.row_country_code,
        )
        .join(UploadedFile, InventoryRow.uploaded_file_id == UploadedFile.id)
        .where(UploadedFile.file_domain == FILE_DOMAIN_SHIPMENT)
    ).all()
    if not row_tuples:
        return

    frame = pd.DataFrame(
        [
            {
                "country_code": r.row_country_code or "KR",
                "base_date": r.row_snapshot_date,
                "sku": str(r.sku or "").strip(),
                "description": r.description or "",
                "supplier": (r.supplier or "").strip(),
                "level": r.level,
                "warehouse": r.warehouse or "",
                "quantity": int(r.quantity or 0),
            }
            for r in row_tuples
            if r.row_snapshot_date is not None
        ]
    )
    if frame.empty:
        return

    grouped = (
        frame.groupby(
            ["country_code", "base_date", "sku", "description", "supplier", "level", "warehouse"],
            dropna=False,
            as_index=False,
        )["quantity"]
        .sum()
    )
    ver = _next_aggregation_version(db)
    now = datetime.now(timezone.utc)
    mappings: list[dict[str, Any]] = []
    for record in grouped.to_dict(orient="records"):
        bd = record["base_date"]
        if bd is None:
            continue
        if not isinstance(bd, date):
            bd = pd.to_datetime(bd).date()
        mappings.append(
            {
                "id": uuid.uuid4(),
                "country_code": str(record["country_code"]),
                "base_date": bd,
                "sku": str(record["sku"] or ""),
                "description": _nullable_string(record["description"]),
                "supplier": _nullable_string(record["supplier"]),
                "level": _nullable_string(record["level"]),
                "warehouse": _nullable_string(record["warehouse"]),
                "total_quantity": int(record["quantity"] or 0),
                "aggregation_version": ver,
                "uploaded_at": now,
                "source_domain": SOURCE_SHIPMENT,
            }
        )
    if mappings:
        db.bulk_insert_mappings(InventoryAggregate, mappings)


def persist_shipment_uploads(
    db: Session,
    file_payloads: list[tuple[UploadFile, bytes, list[dict[str, Any]]]],
) -> list[dict]:
    """이미 파싱된 롱 레코드와 원본 바이트로 S3·inventory_rows 저장(엑셀 재파싱 없음)."""
    settings = get_runtime_settings()
    if not settings.database_enabled:
        return []
    if not is_s3_configured() or not settings.s3_bucket:
        raise HTTPException(status_code=500, detail="S3가 설정되지 않아 출고 파일을 저장할 수 없습니다.")

    s3 = get_s3_client()
    persisted: list[dict] = []
    row_ts = datetime.now(timezone.utc)
    try:
        for upload_file, raw, records in file_payloads:
            if not records:
                continue
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
        ).where(InventoryAggregate.source_domain == SOURCE_SHIPMENT)
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
        if row.base_date is None:
            continue
        long_recs.append(
            {
                "country": row.country_code,
                "channel": row.warehouse or "",
                "sku": row.sku or "",
                "description": row.description or "",
                "brand": row.supplier or "",
                "date": row.base_date,
                "quantity": int(row.total_quantity or 0),
                "level": row.level,
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
