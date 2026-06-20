"""상품 마스터 로우데이터: 엑셀 업로드 → DB 저장 → 그리드 조회·수정."""

from __future__ import annotations

import io
import re
import uuid
from datetime import date, datetime, timezone

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from domains.item.models import ProductGroup, ProductLocale
from domains.item.services.item_mapping import (
    _normalize_brand_cell,
    _normalize_mapping_sku,
    normalize_item_segment,
)
from shared.config import get_runtime_settings

MASTER_COLUMN_ORDER = [
    "brand",
    "segment",
    "version",
    "kr_sku",
    "kr_name",
    "stock_category",
    "fcst_grade",
    "stock_grade",
    "release_month",
    "code_registered_at",
    "us_grade",
    "tw_grade",
    "hk_grade",
    "jp_grade",
]

MASTER_COLUMN_LABELS = {
    "brand": "브랜드",
    "segment": "구분",
    "version": "Ver.",
    "kr_sku": "상품코드",
    "kr_name": "상품명",
    "stock_category": "재고구분",
    "fcst_grade": "FCST등급",
    "stock_grade": "재고등급",
    "release_month": "출시월",
    "code_registered_at": "코드 등록 일자",
    "us_grade": "미국 등급",
    "tw_grade": "대만 등급",
    "hk_grade": "홍콩 등급",
    "jp_grade": "일본 등급",
}

MASTER_HEADER_ALIASES: dict[str, frozenset[str]] = {
    "brand": frozenset({"brand", "브랜드"}),
    "segment": frozenset({"segment", "구분", "상품구분", "분류", "상품분류"}),
    "version": frozenset({"ver", "ver.", "version", "버전", "옵션", "option", "options"}),
    "kr_sku": frozenset({"krsku", "kr_sku", "상품코드", "코드", "sku", "itemno", "item_no"}),
    "kr_name": frozenset({"krname", "kr_name", "상품명", "품명", "description", "name"}),
    "stock_category": frozenset({"stockcategory", "stock_category", "재고구분", "재고분류"}),
    "fcst_grade": frozenset({"fcstgrade", "fcst_grade", "fcst등급", "fcst"}),
    "stock_grade": frozenset({"stockgrade", "stock_grade", "재고등급"}),
    "release_month": frozenset(
        {
            "releasemonth",
            "release_month",
            "출시월",
            "출시일",
            "releasedate",
            "release_date",
            "런칭월",
            "런칭일",
        }
    ),
    "code_registered_at": frozenset(
        {
            "coderegisteredat",
            "code_registered_at",
            "코드등록일자",
            "코드등록일",
            "코드 등록 일자",
            "등록일",
            "등록일자",
        }
    ),
    "us_grade": frozenset({"usgrade", "us_grade", "미국등급", "미국 등급", "us등급"}),
    "tw_grade": frozenset({"twgrade", "tw_grade", "대만등급", "대만 등급", "tw등급"}),
    "hk_grade": frozenset({"hkgrade", "hk_grade", "홍콩등급", "홍콩 등급", "hk등급"}),
    "jp_grade": frozenset({"jpgrade", "jp_grade", "일본등급", "일본 등급", "jp등급"}),
}

MAX_MASTER_FILE_SIZE_BYTES = 5 * 1024 * 1024


def _normalize_master_header(value: object) -> str:
    raw = str(value or "").strip().casefold()
    return re.sub(r"[\s_\-·]+", "", raw)


def _build_master_header_map() -> dict[str, str]:
    out: dict[str, str] = {}
    for canonical, aliases in MASTER_HEADER_ALIASES.items():
        for alias in aliases:
            out[_normalize_master_header(alias)] = canonical
    return out


_MASTER_HEADER_MAP = _build_master_header_map()


def _cell_text(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return str(value).strip()


def _parse_optional_date(value: object, field_label: str, row_label: str) -> date | None:
    raw = _cell_text(value)
    if not raw:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    digits = re.sub(r"[^0-9]", "", raw)
    if len(digits) == 8:
        try:
            return date(int(digits[0:4]), int(digits[4:6]), int(digits[6:8]))
        except ValueError as exc:
            raise HTTPException(
                status_code=400,
                detail=f"{row_label}: {field_label} 날짜 형식이 올바르지 않습니다. (YYYY-MM-DD)",
            ) from exc
    parsed = pd.to_datetime(raw, errors="coerce")
    if pd.isna(parsed):
        raise HTTPException(
            status_code=400,
            detail=f"{row_label}: {field_label} 날짜 형식이 올바르지 않습니다. (YYYY-MM-DD)",
        )
    return parsed.date()


def _parse_release_month(value: object, row_label: str) -> str | None:
    raw = _cell_text(value)
    if not raw:
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y%m")
    if isinstance(value, date):
        return value.strftime("%Y%m")
    digits = re.sub(r"[^0-9]", "", raw)
    if len(digits) == 6:
        year, month = int(digits[0:4]), int(digits[4:6])
        if 1 <= month <= 12:
            return f"{year:04d}{month:02d}"
    if len(digits) == 8:
        year, month = int(digits[0:4]), int(digits[4:6])
        if 1 <= month <= 12:
            return f"{year:04d}{month:02d}"
    parsed = pd.to_datetime(raw, errors="coerce")
    if not pd.isna(parsed):
        return parsed.strftime("%Y%m")
    raise HTTPException(
        status_code=400,
        detail=f"{row_label}: 출시월 형식이 올바르지 않습니다. (YYYYMM)",
    )


def _serialize_optional_date(value: date | None) -> str | None:
    return value.isoformat() if value else None


def _serialize_release_month(group: ProductGroup) -> str:
    if group.release_month:
        return str(group.release_month).strip()
    if group.release_date:
        return group.release_date.strftime("%Y%m")
    return ""


def _kr_locale(group: ProductGroup) -> ProductLocale | None:
    return next(
        (loc for loc in group.locales if str(loc.country_code or "").strip().upper() == "KR"),
        None,
    )


def master_row_payload(group: ProductGroup) -> dict:
    kr = _kr_locale(group)
    seg_norm = normalize_item_segment(group.segment)
    return {
        "group_id": str(group.id),
        "brand": group.brand or "",
        "segment": seg_norm or "",
        "version": group.version or "",
        "kr_sku": kr.sku if kr else "",
        "kr_name": group.kr_name or "",
        "stock_category": group.stock_category or "",
        "fcst_grade": group.fcst_grade or "",
        "stock_grade": group.stock_grade or "",
        "release_month": _serialize_release_month(group),
        "code_registered_at": _serialize_optional_date(group.code_registered_at),
        "us_grade": group.us_grade or "",
        "tw_grade": group.tw_grade or "",
        "hk_grade": group.hk_grade or "",
        "jp_grade": group.jp_grade or "",
        "updated_at": group.updated_at.isoformat() if group.updated_at else None,
    }


def list_item_master_rows(db: Session, query: str | None = None, limit: int = 10000) -> list[dict]:
    groups = (
        db.execute(
            select(ProductGroup)
            .options(selectinload(ProductGroup.locales))
            .order_by(ProductGroup.kr_name.asc())
        )
        .scalars()
        .all()
    )
    needle = str(query or "").strip().casefold()
    cap = max(1, min(limit, 20000))
    rows: list[dict] = []
    for group in groups:
        payload = master_row_payload(group)
        if needle:
            hay = " ".join(str(payload.get(key) or "") for key in MASTER_COLUMN_ORDER).casefold()
            if needle not in hay:
                continue
        rows.append(payload)
        if len(rows) >= cap:
            break
    return rows


def _coerce_master_dataframe(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for col in df.columns:
        canonical = _MASTER_HEADER_MAP.get(_normalize_master_header(col))
        if canonical:
            rename_map[col] = canonical
    if not rename_map:
        raise HTTPException(
            status_code=400,
            detail=f"{filename}: 인식 가능한 헤더가 없습니다. 필수 열: 상품코드, 상품명, 브랜드",
        )
    out = df.rename(columns=rename_map)
    if "kr_sku" not in out.columns or "kr_name" not in out.columns:
        raise HTTPException(
            status_code=400,
            detail=f"{filename}: 상품코드·상품명 열이 필요합니다.",
        )
    return out


def _read_master_records(upload_files: list[UploadFile]) -> list[dict]:
    records: list[dict] = []
    for upload_file in upload_files:
        filename = str(upload_file.filename or "upload.xlsx")
        if not filename.lower().endswith(".xlsx"):
            raise HTTPException(status_code=400, detail=f"{filename}: .xlsx 파일만 업로드할 수 있습니다.")
        upload_file.file.seek(0)
        raw = upload_file.file.read()
        upload_file.file.seek(0)
        if len(raw) > MAX_MASTER_FILE_SIZE_BYTES:
            raise HTTPException(status_code=400, detail=f"{filename}: 파일 크기는 5MB 이하여야 합니다.")
        try:
            df = pd.read_excel(io.BytesIO(raw), sheet_name=0, dtype=object)
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"{filename}: 엑셀을 읽을 수 없습니다. ({exc})") from exc
        df = _coerce_master_dataframe(df, filename)
        for idx, row in df.iterrows():
            row_label = f"{filename} {int(idx) + 2}행"
            kr_sku = _normalize_mapping_sku(row.get("kr_sku"))
            kr_name = _cell_text(row.get("kr_name"))
            if not kr_sku and not kr_name:
                continue
            if not kr_sku:
                raise HTTPException(status_code=400, detail=f"{row_label}: 상품코드가 비어 있습니다.")
            if not kr_name:
                raise HTTPException(status_code=400, detail=f"{row_label}: 상품명이 비어 있습니다.")
            brand = _normalize_brand_cell(row.get("brand")) if "brand" in df.columns else ""
            if not brand:
                raise HTTPException(status_code=400, detail=f"{row_label}: 브랜드는 필수입니다.")
            records.append(
                {
                    "row_label": row_label,
                    "brand": brand,
                    "segment": normalize_item_segment(row.get("segment")) if "segment" in df.columns else None,
                    "version": _cell_text(row.get("version"))[:64] if "version" in df.columns else "",
                    "kr_sku": kr_sku,
                    "kr_name": kr_name,
                    "stock_category": _cell_text(row.get("stock_category"))[:64]
                    if "stock_category" in df.columns
                    else "",
                    "fcst_grade": _cell_text(row.get("fcst_grade"))[:32] if "fcst_grade" in df.columns else "",
                    "stock_grade": _cell_text(row.get("stock_grade"))[:32] if "stock_grade" in df.columns else "",
                    "release_month": _parse_release_month(row.get("release_month"), row_label)
                    if "release_month" in df.columns
                    else None,
                    "code_registered_at": _parse_optional_date(
                        row.get("code_registered_at"), "코드 등록 일자", row_label
                    )
                    if "code_registered_at" in df.columns
                    else None,
                    "us_grade": _cell_text(row.get("us_grade"))[:32] if "us_grade" in df.columns else "",
                    "tw_grade": _cell_text(row.get("tw_grade"))[:32] if "tw_grade" in df.columns else "",
                    "hk_grade": _cell_text(row.get("hk_grade"))[:32] if "hk_grade" in df.columns else "",
                    "jp_grade": _cell_text(row.get("jp_grade"))[:32] if "jp_grade" in df.columns else "",
                }
            )
    if not records:
        raise HTTPException(status_code=400, detail="업로드 파일에서 처리할 데이터 행이 없습니다.")
    return records


def _find_group_by_kr_sku(db: Session, kr_sku_norm: str) -> ProductGroup | None:
    locales = (
        db.execute(
            select(ProductLocale)
            .options(selectinload(ProductLocale.product_group).selectinload(ProductGroup.locales))
            .where(ProductLocale.country_code == "KR")
        )
        .scalars()
        .all()
    )
    for loc in locales:
        if _normalize_mapping_sku(loc.sku) == kr_sku_norm:
            return loc.product_group
    return None


def _apply_master_fields(group: ProductGroup, kr_locale: ProductLocale, record: dict, now: datetime) -> None:
    group.kr_name = record["kr_name"]
    group.brand = record["brand"]
    group.segment = record.get("segment")
    group.version = record.get("version") or None
    group.stock_category = record.get("stock_category") or None
    group.fcst_grade = record.get("fcst_grade") or None
    group.stock_grade = record.get("stock_grade") or None
    group.release_month = record.get("release_month") or None
    group.code_registered_at = record.get("code_registered_at")
    group.us_grade = record.get("us_grade") or None
    group.tw_grade = record.get("tw_grade") or None
    group.hk_grade = record.get("hk_grade") or None
    group.jp_grade = record.get("jp_grade") or None
    kr_locale.sku = record["kr_sku"]
    kr_locale.name = record["kr_name"]
    group.upload_updated_at = now
    group.updated_at = now
    kr_locale.updated_at = now


def merge_item_master_uploads(db: Session, upload_files: list[UploadFile]) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")
    if not upload_files:
        raise HTTPException(status_code=400, detail="업로드할 파일이 필요합니다.")

    records = _read_master_records(upload_files)
    now = datetime.now(timezone.utc)
    touched = 0
    created = 0
    try:
        for record in records:
            existing = _find_group_by_kr_sku(db, record["kr_sku"])
            if existing is not None:
                kr = _kr_locale(existing)
                if kr is None:
                    kr = ProductLocale(
                        id=uuid.uuid4(),
                        item_id=existing.id,
                        country_code="KR",
                        name=record["kr_name"],
                        sku=record["kr_sku"],
                        updated_at=now,
                    )
                    existing.locales.append(kr)
                    db.add(kr)
                _apply_master_fields(existing, kr, record, now)
                touched += 1
                continue

            group = ProductGroup(
                id=uuid.uuid4(),
                kr_name=record["kr_name"],
                brand=record["brand"],
                segment=record.get("segment"),
                version=record.get("version") or None,
                stock_category=record.get("stock_category") or None,
                fcst_grade=record.get("fcst_grade") or None,
                stock_grade=record.get("stock_grade") or None,
                release_month=record.get("release_month"),
                code_registered_at=record.get("code_registered_at"),
                us_grade=record.get("us_grade") or None,
                tw_grade=record.get("tw_grade") or None,
                hk_grade=record.get("hk_grade") or None,
                jp_grade=record.get("jp_grade") or None,
                upload_updated_at=now,
                updated_at=now,
            )
            kr = ProductLocale(
                id=uuid.uuid4(),
                item_id=group.id,
                country_code="KR",
                name=record["kr_name"],
                sku=record["kr_sku"],
                updated_at=now,
            )
            group.locales = [kr]
            db.add(group)
            db.add(kr)
            created += 1
        db.commit()
    except Exception:
        db.rollback()
        raise

    return {
        "processed_file_count": len(upload_files),
        "processed_row_count": len(records),
        "updated_count": touched,
        "created_count": created,
        "columns": [MASTER_COLUMN_LABELS[key] for key in MASTER_COLUMN_ORDER],
    }


def patch_item_master_row(db: Session, payload: dict[str, object]) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")

    raw_gid = str(payload.get("group_id") or "").strip()
    try:
        gid = uuid.UUID(raw_gid)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="유효하지 않은 group_id입니다.") from exc

    group = db.execute(
        select(ProductGroup).where(ProductGroup.id == gid).options(selectinload(ProductGroup.locales))
    ).scalar_one_or_none()
    if group is None:
        raise HTTPException(status_code=404, detail="상품을 찾을 수 없습니다.")

    kr = _kr_locale(group)
    if kr is None:
        kr = ProductLocale(
            id=uuid.uuid4(),
            item_id=group.id,
            country_code="KR",
            name=str(payload.get("kr_name") or group.kr_name or ""),
            sku=str(payload.get("kr_sku") or ""),
            updated_at=datetime.now(timezone.utc),
        )
        group.locales.append(kr)
        db.add(kr)

    kr_sku_in = str(payload.get("kr_sku") or kr.sku or "").strip()
    kr_name_in = str(payload.get("kr_name") or group.kr_name or "").strip()
    brand_in = _normalize_brand_cell(payload.get("brand"))
    if not kr_sku_in:
        raise HTTPException(status_code=400, detail="상품코드는 비울 수 없습니다.")
    if not kr_name_in:
        raise HTTPException(status_code=400, detail="상품명은 비울 수 없습니다.")
    if not brand_in:
        raise HTTPException(status_code=400, detail="브랜드는 비울 수 없습니다.")

    new_norm = _normalize_mapping_sku(kr_sku_in)
    old_norm = _normalize_mapping_sku(kr.sku)
    if new_norm != old_norm:
        conflict = _find_group_by_kr_sku(db, new_norm)
        if conflict is not None and conflict.id != group.id:
            raise HTTPException(status_code=400, detail=f"다른 상품이 이미 사용 중인 상품코드입니다: {kr_sku_in}")

    release_month_in = str(payload.get("release_month") or "").strip()
    record = {
        "brand": brand_in,
        "segment": normalize_item_segment(payload.get("segment"))
        if str(payload.get("segment") or "").strip()
        else None,
        "version": str(payload.get("version") or "").strip()[:64],
        "kr_sku": kr_sku_in,
        "kr_name": kr_name_in,
        "stock_category": str(payload.get("stock_category") or "").strip()[:64],
        "fcst_grade": str(payload.get("fcst_grade") or "").strip()[:32],
        "stock_grade": str(payload.get("stock_grade") or "").strip()[:32],
        "release_month": _parse_release_month(release_month_in, "수정") if release_month_in else None,
        "code_registered_at": _parse_optional_date(
            payload.get("code_registered_at"), "코드 등록 일자", "수정"
        )
        if str(payload.get("code_registered_at") or "").strip()
        else None,
        "us_grade": str(payload.get("us_grade") or "").strip()[:32],
        "tw_grade": str(payload.get("tw_grade") or "").strip()[:32],
        "hk_grade": str(payload.get("hk_grade") or "").strip()[:32],
        "jp_grade": str(payload.get("jp_grade") or "").strip()[:32],
    }
    now = datetime.now(timezone.utc)
    _apply_master_fields(group, kr, record, now)
    group.manual_updated_at = now

    try:
        db.commit()
        fresh = db.execute(
            select(ProductGroup).where(ProductGroup.id == gid).options(selectinload(ProductGroup.locales))
        ).scalar_one()
    except Exception:
        db.rollback()
        raise
    return master_row_payload(fresh)


__all__ = [
    "MASTER_COLUMN_LABELS",
    "MASTER_COLUMN_ORDER",
    "list_item_master_rows",
    "merge_item_master_uploads",
    "master_row_payload",
    "patch_item_master_row",
]
