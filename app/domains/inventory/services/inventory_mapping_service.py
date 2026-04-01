import io
import re
from datetime import datetime, timezone

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import delete, func, select
from sqlalchemy.orm import Session

from app.domains.inventory.models import ProductSkuMapping
from app.domains.inventory.services.inventory_aggregate_service import _normalize_item_code
from app.shared.config import get_runtime_settings

MAPPING_TEMPLATE_COLUMNS = ["description", "kr", "us", "tw", "vn", "sg", "au", "uk", "ae"]
MAPPING_SKU_COLUMNS = [column for column in MAPPING_TEMPLATE_COLUMNS if column != "description"]
COUNTRY_TO_MAPPING_COLUMN = {
    "KR": "kr",
    "US": "us",
    "TW": "tw",
    "VN": "vn",
    "SG": "sg",
    "AU": "au",
    "UK": "uk",
    "AE": "ae",
}
MAX_MAPPING_FILE_SIZE_BYTES = 2 * 1024 * 1024


def _normalize_mapping_description(value: object) -> str:
    raw = str(value or "").strip()
    return re.sub(r"\s+", " ", raw)


def _normalize_mapping_description_key(value: object) -> str:
    return _normalize_mapping_description(value).casefold()


def _normalize_mapping_sku(value: object) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    return _normalize_item_code(raw)


def _build_mapping_record(source: dict[str, object]) -> dict[str, str]:
    return {
        "description": _normalize_mapping_description(source.get("description")),
        "kr": _normalize_mapping_sku(source.get("kr")),
        "us": _normalize_mapping_sku(source.get("us")),
        "tw": _normalize_mapping_sku(source.get("tw")),
        "vn": _normalize_mapping_sku(source.get("vn")),
        "sg": _normalize_mapping_sku(source.get("sg")),
        "au": _normalize_mapping_sku(source.get("au")),
        "uk": _normalize_mapping_sku(source.get("uk")),
        "ae": _normalize_mapping_sku(source.get("ae")),
    }


def _validate_mapping_record(record: dict[str, str], row_label: str) -> dict[str, str]:
    if not record["description"]:
        raise HTTPException(status_code=400, detail=f"{row_label}: description은 필수입니다.")
    if not record["kr"]:
        raise HTTPException(status_code=400, detail=f"{row_label}: kr은 필수입니다.")
    return record


def _mapping_item_payload(row: ProductSkuMapping) -> dict:
    return {
        "description": row.description,
        "kr": row.kr,
        "us": row.us,
        "tw": row.tw,
        "vn": row.vn,
        "sg": row.sg,
        "au": row.au,
        "uk": row.uk,
        "ae": row.ae,
    }


def _find_mapping_conflict(
    rows: list[ProductSkuMapping],
    field_name: str,
    field_value: str,
    exclude_id: object = None,
) -> ProductSkuMapping | None:
    if not field_value:
        return None
    for row in rows:
        if exclude_id is not None and row.id == exclude_id:
            continue
        current_value = str(getattr(row, field_name) or "").strip()
        if current_value == field_value:
            return row
    return None


def _read_mapping_bytes(upload_file: UploadFile) -> tuple[bytes, int]:
    filename = str(upload_file.filename or "mapping.xlsx")
    if not filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="매핑 마스터는 .xlsx 파일만 업로드할 수 있습니다.")

    upload_file.file.seek(0)
    raw = upload_file.file.read()
    upload_file.file.seek(0)
    size = len(raw or b"")
    if size == 0:
        raise HTTPException(status_code=400, detail=f"{filename}: 파일이 비어 있습니다.")
    if size > MAX_MAPPING_FILE_SIZE_BYTES:
        raise HTTPException(
            status_code=400,
            detail=f"{filename}: 파일 크기는 {MAX_MAPPING_FILE_SIZE_BYTES // (1024 * 1024)}MB 이하여야 합니다.",
        )
    return raw, size


def _read_mapping_frame(upload_file: UploadFile) -> pd.DataFrame:
    raw, _ = _read_mapping_bytes(upload_file)
    filename = str(upload_file.filename or "mapping.xlsx")
    try:
        df = pd.read_excel(io.BytesIO(raw), sheet_name=0, dtype=object)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"{filename}: 엑셀 파싱 실패 ({exc})") from exc

    actual_columns = [str(column).strip() for column in df.columns]
    if actual_columns != MAPPING_TEMPLATE_COLUMNS:
        raise HTTPException(
            status_code=400,
            detail=(
                f"{filename}: 템플릿 헤더가 올바르지 않습니다. "
                f"필수 헤더 순서: {', '.join(MAPPING_TEMPLATE_COLUMNS)}"
            ),
        )
    if df.empty:
        raise HTTPException(status_code=400, detail=f"{filename}: 최소 1개 이상의 매핑 행이 필요합니다.")
    return df


def _validate_and_normalize_mapping_frame(df: pd.DataFrame) -> list[dict]:
    normalized_records: list[dict] = []
    seen_name_rows: dict[str, int] = {}
    seen_sku_rows: dict[str, dict[str, int]] = {column: {} for column in MAPPING_SKU_COLUMNS}

    for idx, row in df.iterrows():
        excel_row = idx + 2
        record = _validate_mapping_record(_build_mapping_record(row.to_dict()), f"행 {excel_row}")

        description_key = _normalize_mapping_description_key(record["description"])
        prior_name_row = seen_name_rows.get(description_key)
        if prior_name_row is not None:
            raise HTTPException(
                status_code=400,
                detail=f"행 {excel_row}: description이 중복됩니다. (기존 행 {prior_name_row})",
            )
        seen_name_rows[description_key] = excel_row

        for column in MAPPING_SKU_COLUMNS:
            sku = str(record[column] or "").strip()
            if not sku:
                continue
            prior_sku_row = seen_sku_rows[column].get(sku)
            if prior_sku_row is not None:
                raise HTTPException(
                    status_code=400,
                    detail=f"행 {excel_row}: {column} 값 `{sku}` 가 중복됩니다. (기존 행 {prior_sku_row})",
                )
            seen_sku_rows[column][sku] = excel_row

        normalized_records.append(record)

    return normalized_records


def get_product_sku_mapping_summary(db: Session) -> dict:
    total_count = db.execute(select(func.count(ProductSkuMapping.id))).scalar_one()
    updated_at = db.execute(select(func.max(ProductSkuMapping.updated_at))).scalar_one()
    upload_updated_at = db.execute(select(func.max(ProductSkuMapping.upload_updated_at))).scalar_one()
    manual_updated_at = db.execute(select(func.max(ProductSkuMapping.manual_updated_at))).scalar_one()
    return {
        "total_count": int(total_count or 0),
        "updated_at": updated_at.isoformat() if updated_at else None,
        "upload_updated_at": upload_updated_at.isoformat() if upload_updated_at else None,
        "manual_updated_at": manual_updated_at.isoformat() if manual_updated_at else None,
        "required_columns": MAPPING_TEMPLATE_COLUMNS,
    }


def replace_product_sku_mappings(db: Session, upload_file: UploadFile) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")

    df = _read_mapping_frame(upload_file)
    records = _validate_and_normalize_mapping_frame(df)
    uploaded_at = datetime.now(timezone.utc)

    try:
        db.execute(delete(ProductSkuMapping))
        db.flush()
        for record in records:
            db.add(ProductSkuMapping(**record, upload_updated_at=uploaded_at, manual_updated_at=None))
        db.commit()
    except Exception:
        db.rollback()
        raise

    summary = get_product_sku_mapping_summary(db)
    summary["replaced_count"] = len(records)
    return summary


def clear_product_sku_mappings(db: Session) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")

    deleted_count = db.execute(select(func.count(ProductSkuMapping.id))).scalar_one()
    try:
        db.execute(delete(ProductSkuMapping))
        db.commit()
    except Exception:
        db.rollback()
        raise
    return {"deleted_count": int(deleted_count or 0)}


def upsert_product_sku_mapping(db: Session, payload: dict[str, object]) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")

    record = _validate_mapping_record(_build_mapping_record(payload), "수기 입력")
    rows = db.execute(select(ProductSkuMapping).order_by(ProductSkuMapping.description.asc())).scalars().all()
    manual_updated_at = datetime.now(timezone.utc)

    description_conflict = _find_mapping_conflict(rows, "description", record["description"])
    kr_conflict = _find_mapping_conflict(rows, "kr", record["kr"])
    if description_conflict and kr_conflict and description_conflict.id != kr_conflict.id:
        raise HTTPException(
            status_code=400,
            detail="description과 kr이 서로 다른 기존 매핑을 가리키고 있습니다. 둘 중 하나를 기준으로 먼저 정리해주세요.",
        )

    target = description_conflict or kr_conflict
    target_id = target.id if target else None
    for column in MAPPING_TEMPLATE_COLUMNS:
        value = str(record.get(column) or "").strip()
        conflict = _find_mapping_conflict(rows, column, value, exclude_id=target_id)
        if conflict is not None:
            raise HTTPException(
                status_code=400,
                detail=f"`{column}` 값 `{value}` 는 이미 다른 SKU 매핑에 등록되어 있습니다.",
            )

    try:
        if target is None:
            target = ProductSkuMapping(**record, manual_updated_at=manual_updated_at)
            db.add(target)
        else:
            for column, value in record.items():
                normalized_value = value
                if column not in {"description", "kr"} and not normalized_value:
                    normalized_value = None
                setattr(target, column, normalized_value)
            target.manual_updated_at = manual_updated_at
        db.commit()
        db.refresh(target)
    except Exception:
        db.rollback()
        raise

    return _mapping_item_payload(target)


def list_product_sku_mappings(db: Session, query: str | None = None, limit: int = 100) -> list[dict]:
    rows = db.execute(select(ProductSkuMapping).order_by(ProductSkuMapping.description.asc())).scalars().all()
    needle = str(query or "").strip().casefold()
    items: list[dict] = []
    for row in rows:
        item = _mapping_item_payload(row)
        if needle:
            haystack = " ".join(str(value or "") for value in item.values()).casefold()
            if needle not in haystack:
                continue
        items.append(item)
        if len(items) >= max(1, limit):
            break
    return items


def build_product_sku_mapping_lookups(db: Session) -> dict:
    rows = db.execute(select(ProductSkuMapping).order_by(ProductSkuMapping.description.asc())).scalars().all()
    by_country_sku: dict[str, dict[str, dict[str, str]]] = {code: {} for code in COUNTRY_TO_MAPPING_COLUMN}
    by_description_key: dict[str, dict[str, str]] = {}

    for row in rows:
        payload = {"description": row.description, "kr": row.kr}
        by_description_key[_normalize_mapping_description_key(row.description)] = payload
        for country_code, column_name in COUNTRY_TO_MAPPING_COLUMN.items():
            sku = str(getattr(row, column_name) or "").strip()
            if not sku:
                continue
            by_country_sku[country_code][sku] = payload

    return {"by_country_sku": by_country_sku, "by_description_key": by_description_key}


def resolve_product_sku_mapping(
    lookups: dict,
    country_code: str,
    raw_item_code: object,
    description: object = "",
) -> dict[str, str]:
    normalized_country = str(country_code or "").strip().upper() or "KR"
    normalized_sku = _normalize_mapping_sku(raw_item_code)
    country_rows = ((lookups or {}).get("by_country_sku") or {}).get(normalized_country, {})
    matched = country_rows.get(normalized_sku) if normalized_sku else None
    if not matched and normalized_country == "KR":
        matched = ((lookups or {}).get("by_description_key") or {}).get(_normalize_mapping_description_key(description))
    if not matched:
        return {"mapped_kr_name": "", "mapped_kr_sku": ""}

    return {
        "mapped_kr_name": str(matched.get("description") or ""),
        "mapped_kr_sku": str(matched.get("kr") or ""),
    }
