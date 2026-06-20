"""출고 ?�일 S3·DB ?�??�?집계 조회.

POST /api/inventory/shipment/aggregate ??**?�시 4??*�?처리?�다.
(?�매처·상?�코???�는 ?�드민코?�·주문일·주문?�량). ?�트명·파?�명?� ?�용?��? ?�으�?
�??�트부???�더가 맞는 ?�트�?찾는?? ?�매처→칩·통??집계??
`shipment_vendor_channel_maps` ??1�??�합 규칙???�른??
"""

from __future__ import annotations

import io
import os
import uuid
from datetime import datetime, timezone
from typing import Any

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import delete, func, literal, select
from sqlalchemy.orm import Session

from domains.shipment.models import ShipmentMatrixRow
from domains.stock.models import InventoryAggregate, InventoryRow, UploadedFile
from domains.stock.services.stock_persistence import (
    _next_aggregation_version,
    _nullable_string,
    _serialize_uploaded_file,
)
from domains.shipment.services.shipment_matrix import (
    DEFAULT_SHIPMENT_MATRIX_YEAR,
    build_shipment_matrix_view,
    list_shipment_matrix_years,
    parse_shipment_raw_orders_workbook,
    shipment_workbook_is_raw_orders_format,
)
from shared.config import get_runtime_settings
from shared.storage import get_s3_client, is_s3_configured

SOURCE_SHIPMENT = "SHIPMENT"
FILE_DOMAIN_SHIPMENT = "shipment"
SHIPMENT_BUCKET_COUNTRY = "SHIPMENT"
SHIPMENT_MATRIX_UPLOAD_KEY = "matrix"


def _shipment_s3_key(prefix: str, stored_name: str) -> str:
    root = (prefix or "inventory").strip().strip("/")
    return f"{root}/shipment/{stored_name}"


def purge_shipment_matrix_rows_if_no_matrix_uploads(db: Session) -> None:
    """`shipment_sheet_key == matrix` ???�로?��? ?�나???�으�?`shipment_matrix_rows` ?��? ??��.

    - ?�시 4???�로?�도 ???��? ?�.
    - ?�전 ?�이?�는 `file_domain`??`inventory`�??�아 ??�� ??출고 분기�??�지 ?�는 경우가 ?�어,
      ?�출부?�서 `country_code == SHIPMENT` ??��??출고 ?�리�?묶음.
    - `file_domain` 조건?� ?��? ?�음(?�메???�락 ?�거?��? ?�일?�게 matrix ?�만 보면 ??.
    """
    n = db.scalar(
        select(func.count()).select_from(UploadedFile).where(
            UploadedFile.shipment_sheet_key == SHIPMENT_MATRIX_UPLOAD_KEY,
        )
    )
    if (n or 0) == 0:
        db.execute(delete(ShipmentMatrixRow))


def _rebuild_all_shipment_aggregates(db: Session) -> None:
    """출고 inventory_rows ?�량??메모리에 ?�리지 ?�고 DB?�서 GROUP BY ??집계 ?�만 ?�재."""
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


def _upsert_shipment_matrix_channel_rows(
    db: Session,
    channel_sheet: str,
    data_year: int,
    recs: list[dict[str, Any]],
) -> None:
    """같�? ?�도 ?�업로드: ?�일???�는 (채널×SKU)�??�·일 ?�??병합(�?변�???갱신, ????추�?). ?�일???�는 SKU·?�?� ?��?."""
    recs_norm: list[dict[str, Any]] = []
    sku_keys: list[str] = []
    for rec in recs:
        sku = str(rec.get("sku") or "").strip()
        if not sku:
            continue
        recs_norm.append(rec)
        sku_keys.append(sku)
    if not recs_norm:
        return

    # 병목 ?�거: SKU�?N�?SELECT ?�??채널/?�도 기�??�로 ??번에 기존 ??조회
    existing_rows = (
        db.execute(
            select(ShipmentMatrixRow).where(
                ShipmentMatrixRow.channel_sheet == channel_sheet,
                ShipmentMatrixRow.data_year == data_year,
                ShipmentMatrixRow.sku.in_(sku_keys),
            )
        )
        .scalars()
        .all()
    )
    existing_by_sku = {str(row.sku): row for row in existing_rows}

    for rec in recs_norm:
        sku = str(rec.get("sku") or "").strip()
        mt_new = {str(k): int(v) for k, v in (rec.get("monthly") or {}).items()}
        dt_new = {str(k): int(v) for k, v in (rec.get("daily") or {}).items()}
        mkt_raw = rec.get("mkt_priority")
        cat_raw = rec.get("category")
        mkt = str(mkt_raw).strip() if mkt_raw is not None and str(mkt_raw).strip() else None
        cat = str(cat_raw).strip() if cat_raw is not None and str(cat_raw).strip() else None

        existing = existing_by_sku.get(sku)
        if existing:
            mt = {str(k): int(v) for k, v in (existing.monthly_totals or {}).items()}
            for k, v in mt_new.items():
                mt[k] = v
            dt = {str(k): int(v) for k, v in (existing.daily_totals or {}).items()}
            for k, v in dt_new.items():
                dt[k] = v
            existing.monthly_totals = mt or None
            existing.daily_totals = dt or None
            if mkt is not None:
                existing.mkt_priority = mkt
            if cat is not None:
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
    kr_sku_brand_name: dict[str, tuple[str, str, str, str]] | None = None,
    vendor_primary_overrides_internal: dict[str, str] | None = None,
) -> tuple[list[dict], list[int]]:
    """?�시 출고 ?��?(?�매처·상?�코?�·주문일·주문?�량): 주문?�에???�도 추론 ??칩별 ?�서?�·S3 ?�?? ?�일명·시?�명 ?�약 ?�음."""
    settings = get_runtime_settings()
    if not settings.database_enabled:
        return [], []
    if not is_s3_configured() or not settings.s3_bucket:
        raise HTTPException(status_code=500, detail="S3가 ?�정?��? ?�아 출고 ?�일???�?�할 ???�습?�다.")

    s3 = get_s3_client()
    persisted: list[dict] = []
    years_order: list[int] = []
    try:
        for upload_file, raw in files:
            bio = io.BytesIO(raw)
            xl = pd.ExcelFile(bio)
            if shipment_workbook_is_raw_orders_format(xl):
                parsed_by_year, years_from_file = parse_shipment_raw_orders_workbook(
                    db,
                    raw,
                    upload_file.filename or "",
                    kr_sku_brand_name=kr_sku_brand_name,
                    vendor_primary_overrides_internal=vendor_primary_overrides_internal,
                )
            else:
                raise HTTPException(
                    status_code=400,
                    detail={
                        "code": "SHIPMENT_WRONG_FORMAT",
                        "message": (
                            f"{upload_file.filename or '?�일'}: 출고 ?�식???????�습?�다. "
                            "?�트 �??�나???�판매처?�「상?�코?��??�는 ?�어?��?코드???�어?��? 코드???�주문일?�「주문수?��??�는 ?�주�??��? ?�이 ?�요?�니??"
                        ),
                        "filename": (upload_file.filename or "").strip(),
                        "sheet_name": None,
                    },
                )
            years_order.extend(years_from_file)
            for data_year in sorted(parsed_by_year.keys()):
                parsed = parsed_by_year[data_year]
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
        return persisted, years_order
    except Exception:
        db.rollback()
        raise


def get_shipment_view(
    db: Session,
    *,
    channel: str | None = None,
    data_year: int = DEFAULT_SHIPMENT_MATRIX_YEAR,
    all_channels: bool = False,
) -> dict:
    """출고 ?�황 매트�?��(�??�트�?. �??�에 ?�계 ???�함."""
    payload = build_shipment_matrix_view(
        db, channel=channel, data_year=data_year, all_channels=all_channels
    )
    totals_row = payload.pop("shipment_totals_row", None)
    body_rows: list[dict[str, Any]] = list(payload.get("rows") or [])
    if totals_row:
        payload["rows"] = [totals_row, *body_rows]
    else:
        payload["rows"] = body_rows
    payload["shipment_available_years"] = list_shipment_matrix_years(db)
    payload["shipment_matrix_year"] = data_year
    return payload
