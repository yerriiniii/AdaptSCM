"""출고 파일 S3·DB 저장 및 집계 조회.

POST /api/inventory/shipment/aggregate 는 **원시 4열**만 처리한다.
(판매처·상품코드 또는 어드민코드·주문일·주문수량). 시트명·파일명은 사용하지 않으며,
첫 시트부터 헤더가 맞는 시트를 찾는다. 판매처→칩·통합 집계는
`shipment_vendor_channel_maps` 의 1차/통합 규칙을 따른다.
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

from domains.inventory.models import InventoryAggregate, InventoryRow, ShipmentMatrixRow, UploadedFile
from domains.inventory.services.inventory_persistence_service import (
    _next_aggregation_version,
    _nullable_string,
    _serialize_uploaded_file,
)
from domains.inventory.services.shipment_matrix_service import (
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
    """`shipment_sheet_key == matrix` 인 업로드가 하나도 없으면 `shipment_matrix_rows` 전부 삭제.

    - 원시 4열 업로드도 이 키를 씀.
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


def _upsert_shipment_matrix_channel_rows(
    db: Session,
    channel_sheet: str,
    data_year: int,
    recs: list[dict[str, Any]],
) -> None:
    """같은 연도 재업로드: 파일에 있는 (채널×SKU)만 월·일 셀을 병합(값 변경 시 갱신, 새 키 추가). 파일에 없는 SKU·셀은 유지."""
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
    """원시 출고 엑셀(판매처·상품코드·주문일·주문수량): 주문일에서 연도 추론 후 칩별 업서트·S3 저장. 파일명·시트명 제약 없음."""
    settings = get_runtime_settings()
    if not settings.database_enabled:
        return [], []
    if not is_s3_configured() or not settings.s3_bucket:
        raise HTTPException(status_code=500, detail="S3가 설정되지 않아 출고 파일을 저장할 수 없습니다.")

    s3 = get_s3_client()
    persisted: list[dict] = []
    years_order: list[int] = []
    try:
        for upload_file, raw in files:
            bio = io.BytesIO(raw)
            xl = pd.ExcelFile(bio)
            if shipment_workbook_is_raw_orders_format(xl):
                parsed, data_year = parse_shipment_raw_orders_workbook(
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
                            f"{upload_file.filename or '파일'}: 출고 형식을 알 수 없습니다. "
                            "시트 중 하나에 「판매처」「상품코드」(또는 「어드민코드」/「어드민 코드」)「주문일」「주문수량」(또는 「주문 수」) 열이 필요합니다."
                        ),
                        "filename": (upload_file.filename or "").strip(),
                        "sheet_name": None,
                    },
                )
            years_order.append(data_year)
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
    """출고 현황 매트릭스(칩=시트명). 첫 행에 합계 행 포함."""
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
