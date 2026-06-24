"""상품 마스터 로우데이터: 엑셀 업로드 → DB 저장 → 그리드 조회·수정."""

from __future__ import annotations

import io
import re
import uuid
from datetime import date, datetime, timezone

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import inspect as sa_inspect
from sqlalchemy import select
from sqlalchemy.exc import IntegrityError
from sqlalchemy.orm import Session, selectinload

from domains.item.models import ItemMasterExtraColumn, ProductGroup, ProductLocale
from domains.item.services.item_mapping import (
    _normalize_brand_cell,
    _normalize_barcode_cell,
    _normalize_mapping_sku,
    normalize_item_segment,
)
from shared.config import get_runtime_settings

MASTER_COLUMN_ORDER = [
    "brand",
    "representative_code",
    "kr_sku",
    "version",
    "kr_name",
    "segment",
    "stock_category",
    "fcst_grade",
    "stock_grade",
    "release_month",
    "code_registered_at",
    "kr_grade",
    "us_grade",
    "tw_grade",
    "hk_grade",
    "jp_grade",
]

MASTER_COLUMN_LABELS = {
    "brand": "브랜드",
    "representative_code": "대표코드",
    "kr_sku": "상품코드",
    "version": "Ver.",
    "kr_name": "상품명",
    "segment": "구분",
    "stock_category": "재고구분",
    "fcst_grade": "FCST등급",
    "stock_grade": "재고등급",
    "release_month": "출시월",
    "code_registered_at": "코드 등록 일자",
    "kr_grade": "한국 등급",
    "us_grade": "미국 등급",
    "tw_grade": "대만 등급",
    "hk_grade": "홍콩 등급",
    "jp_grade": "일본 등급",
}

MASTER_HEADER_ALIASES: dict[str, frozenset[str]] = {
    "brand": frozenset({"brand", "브랜드"}),
    "representative_code": frozenset(
        {"representativecode", "representative_code", "대표코드", "repcode", "rep_code"}
    ),
    "segment": frozenset({"segment", "구분", "상품구분", "분류", "상품분류"}),
    "version": frozenset({"ver", "ver.", "version", "버전", "옵션", "option", "options"}),
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
    "kr_grade": frozenset({"krgrade", "kr_grade", "한국등급", "한국 등급", "kr등급"}),
    "tw_grade": frozenset({"twgrade", "tw_grade", "대만등급", "대만 등급", "tw등급"}),
    "hk_grade": frozenset({"hkgrade", "hk_grade", "홍콩등급", "홍콩 등급", "hk등급"}),
    "jp_grade": frozenset({"jpgrade", "jp_grade", "일본등급", "일본 등급", "jp등급"}),
}

MAX_MASTER_FILE_SIZE_BYTES = 5 * 1024 * 1024
MAX_EXTRA_COLUMNS = 40
EXTRA_COLUMN_LABEL_MAX_LEN = 64
EXTRA_FIELD_VALUE_MAX_LEN = 255


def _normalize_master_header(value: object) -> str:
    raw = str(value or "").strip().casefold()
    return re.sub(r"[\s_\-·]+", "", raw)


# 상품코드 열: 이 헤더만 인식 (우선순위 — 앞일수록 우선). 대표코드는 별도 열.
KR_SKU_HEADER_LABELS: tuple[str, ...] = ("상품코드", "어드민코드")
KR_SKU_HEADER_NORMS: frozenset[str] = frozenset(
    _normalize_master_header(label) for label in KR_SKU_HEADER_LABELS
)


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


RELEASE_MONTH_MAX_LEN = 64
CODE_REGISTERED_AT_MAX_LEN = 64


def _normalize_release_month(value: object) -> str | None:
    """출시월 — YYYYMM 고정이 아니라 자유 텍스트."""
    raw = _cell_text(value)
    if not raw:
        return None
    return raw[:RELEASE_MONTH_MAX_LEN]


def _normalize_code_registered_at(value: object) -> str | None:
    """코드 등록 일자 — 날짜 파싱 없이 자유 텍스트."""
    raw = _cell_text(value)
    if not raw:
        return None
    return raw[:CODE_REGISTERED_AT_MAX_LEN]


def _serialize_code_registered_at(group: ProductGroup) -> str:
    raw = _cell_text(group.code_registered_at)
    if not raw:
        return ""
    if "T" in raw:
        return raw.split("T", 1)[0]
    if " " in raw and len(raw) >= 10 and raw[4:5] == "-" and raw[7:8] == "-":
        return raw.split(" ", 1)[0]
    return raw


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


def _extra_fields_dict(raw: object) -> dict:
    if not isinstance(raw, dict):
        return {}
    return raw


def list_item_master_extra_columns(db: Session) -> list[dict]:
    rows = (
        db.execute(
            select(ItemMasterExtraColumn).order_by(
                ItemMasterExtraColumn.sort_order.asc(),
                ItemMasterExtraColumn.created_at.asc(),
            )
        )
        .scalars()
        .all()
    )
    return [
        {
            "field_key": row.field_key,
            "label": row.label,
            "sort_order": row.sort_order,
        }
        for row in rows
    ]


def _slug_extra_field_key(label: str) -> str:
    raw = re.sub(r"\s+", "_", str(label or "").strip())[:48]
    raw = re.sub(r"[^\w가-힣]", "", raw, flags=re.UNICODE)
    if not raw:
        raw = "col"
    return f"x_{raw}"[:64]


def add_item_master_extra_column(db: Session, label: str) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")
    label_in = str(label or "").strip()[:EXTRA_COLUMN_LABEL_MAX_LEN]
    if not label_in:
        raise HTTPException(status_code=400, detail="열 이름을 입력해 주세요.")
    existing = list_item_master_extra_columns(db)
    if len(existing) >= MAX_EXTRA_COLUMNS:
        raise HTTPException(status_code=400, detail=f"사용자 정의 열은 최대 {MAX_EXTRA_COLUMNS}개까지 추가할 수 있습니다.")
    if any(str(c["label"]).strip() == label_in for c in existing):
        raise HTTPException(status_code=400, detail=f"이미 같은 이름의 열이 있습니다: {label_in}")
    base_key = _slug_extra_field_key(label_in)
    field_key = base_key
    used = {c["field_key"] for c in existing}
    n = 2
    while field_key in used:
        suffix = f"_{n}"
        field_key = f"{base_key[: 64 - len(suffix)]}{suffix}"
        n += 1
    sort_order = max((int(c["sort_order"]) for c in existing), default=-1) + 1
    row = ItemMasterExtraColumn(
        id=uuid.uuid4(),
        field_key=field_key,
        label=label_in,
        sort_order=sort_order,
        created_at=datetime.now(timezone.utc),
    )
    db.add(row)
    try:
        db.commit()
        db.refresh(row)
    except Exception:
        db.rollback()
        raise
    return {"field_key": row.field_key, "label": row.label, "sort_order": row.sort_order}


def delete_item_master_extra_column(db: Session, field_key: str) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")
    key = str(field_key or "").strip()
    if not key:
        raise HTTPException(status_code=400, detail="field_key가 필요합니다.")
    row = db.execute(
        select(ItemMasterExtraColumn).where(ItemMasterExtraColumn.field_key == key)
    ).scalar_one_or_none()
    if row is None:
        raise HTTPException(status_code=404, detail="삭제할 열을 찾을 수 없습니다.")
    label = row.label
    groups = (
        db.execute(select(ProductGroup).where(ProductGroup.extra_fields.isnot(None))).scalars().all()
    )
    for group in groups:
        raw = _extra_fields_dict(group.extra_fields)
        if key not in raw:
            continue
        merged = dict(raw)
        del merged[key]
        group.extra_fields = merged or None
    db.delete(row)
    try:
        db.commit()
    except Exception:
        db.rollback()
        raise
    return {"field_key": key, "label": label, "deleted": True}


def _extra_fields_payload(group: ProductGroup, field_keys: list[str]) -> dict[str, str]:
    raw = _extra_fields_dict(group.extra_fields)
    return {key: _cell_text(raw.get(key))[:EXTRA_FIELD_VALUE_MAX_LEN] for key in field_keys}


def _resolve_representative_code(explicit: object, kr_sku: str) -> str:
    raw = _normalize_mapping_sku(explicit)
    if raw:
        return raw
    return kr_sku


def master_row_payload(group: ProductGroup, extra_field_keys: list[str] | None = None) -> dict:
    kr = _kr_locale(group)
    seg_norm = normalize_item_segment(group.segment)
    kr_sku = kr.sku if kr else ""
    rep = _cell_text(group.representative_code) or kr_sku
    return {
        "group_id": str(group.id),
        "brand": group.brand or "",
        "representative_code": rep,
        "segment": seg_norm or "",
        "version": group.version or "",
        "kr_sku": kr_sku,
        "kr_name": group.kr_name or "",
        "stock_category": group.stock_category or "",
        "fcst_grade": group.fcst_grade or "",
        "stock_grade": group.stock_grade or "",
        "release_month": _serialize_release_month(group),
        "code_registered_at": _serialize_code_registered_at(group),
        "kr_grade": group.kr_grade or "",
        "us_grade": group.us_grade or "",
        "tw_grade": group.tw_grade or "",
        "hk_grade": group.hk_grade or "",
        "jp_grade": group.jp_grade or "",
        "barcode": group.barcode or "",
        "extra_fields": _extra_fields_payload(group, extra_field_keys or []),
        "updated_at": group.updated_at.isoformat() if group.updated_at else None,
    }


def get_item_master_row_by_group_id(db: Session, group_id: str) -> tuple[dict, list[dict]]:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")
    gid_raw = str(group_id or "").strip()
    if not gid_raw:
        raise HTTPException(status_code=400, detail="group_id가 필요합니다.")
    try:
        gid = uuid.UUID(gid_raw)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="유효하지 않은 group_id입니다.") from exc
    extra_cols = list_item_master_extra_columns(db)
    extra_keys = [str(c["field_key"]) for c in extra_cols]
    group = db.execute(
        select(ProductGroup)
        .where(ProductGroup.id == gid)
        .options(selectinload(ProductGroup.locales))
    ).scalar_one_or_none()
    if group is None:
        raise HTTPException(status_code=404, detail="상품을 찾을 수 없습니다.")
    return master_row_payload(group, extra_keys), extra_cols


def list_item_master_rows(
    db: Session,
    query: str | None = None,
    limit: int = 10000,
    extra_columns: list[dict] | None = None,
) -> tuple[list[dict], list[dict]]:
    extra_cols = extra_columns if extra_columns is not None else list_item_master_extra_columns(db)
    extra_keys = [str(c["field_key"]) for c in extra_cols]
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
        payload = master_row_payload(group, extra_keys)
        if needle:
            hay_parts = [str(payload.get(key) or "") for key in MASTER_COLUMN_ORDER]
            hay_parts.extend(str(payload.get("extra_fields", {}).get(key) or "") for key in extra_keys)
            hay = " ".join(hay_parts).casefold()
            if needle not in hay:
                continue
        rows.append(payload)
        if len(rows) >= cap:
            break
    return rows, extra_cols


def _kr_sku_source_columns(df: pd.DataFrame) -> list[str]:
    """원본 시트 열 중 상품코드 후보 (KR_SKU_HEADER_LABELS 순)."""
    col_by_norm: dict[str, str] = {}
    for col in df.columns:
        norm = _normalize_master_header(col)
        if norm in KR_SKU_HEADER_NORMS and norm not in col_by_norm:
            col_by_norm[norm] = col
    return [col_by_norm[n] for n in (_normalize_master_header(l) for l in KR_SKU_HEADER_LABELS) if n in col_by_norm]


def _merge_kr_sku_series(df: pd.DataFrame) -> pd.Series:
    """상품코드 > 어드민코드 — 첫 번째 비어 있지 않은 값."""
    col_by_norm: dict[str, str] = {}
    for col in df.columns:
        norm = _normalize_master_header(col)
        if norm in KR_SKU_HEADER_NORMS and norm not in col_by_norm:
            col_by_norm[norm] = col
    priority_norms = [_normalize_master_header(label) for label in KR_SKU_HEADER_LABELS]

    def pick(row: pd.Series) -> str:
        for norm in priority_norms:
            col = col_by_norm.get(norm)
            if col is None:
                continue
            sku = _normalize_mapping_sku(row[col])
            if sku:
                return sku
        return ""

    return df.apply(pick, axis=1)


def _coerce_master_dataframe(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for col in df.columns:
        norm = _normalize_master_header(col)
        if norm in KR_SKU_HEADER_NORMS:
            continue
        canonical = _MASTER_HEADER_MAP.get(norm)
        if canonical:
            rename_map[col] = canonical
    if not rename_map and not _kr_sku_source_columns(df):
        raise HTTPException(
            status_code=400,
            detail=f"{filename}: 인식 가능한 헤더가 없습니다. 필수 열: 상품코드, 상품명, 브랜드",
        )
    out = df.rename(columns=rename_map)
    sku_sources = _kr_sku_source_columns(df)
    if sku_sources:
        out["kr_sku"] = _merge_kr_sku_series(df)
    if "kr_sku" not in out.columns or "kr_name" not in out.columns:
        raise HTTPException(
            status_code=400,
            detail=(
                f"{filename}: 상품코드·상품명 열이 필요합니다. "
                f"상품코드는 「{'」「'.join(KR_SKU_HEADER_LABELS)}」 헤더만 사용합니다."
            ),
        )
    return out


def _master_header_matches_in_row(row_values: list[object]) -> set[str]:
    matched: set[str] = set()
    for cell in row_values:
        norm = _normalize_master_header(cell)
        if norm in KR_SKU_HEADER_NORMS:
            matched.add("kr_sku")
            continue
        canonical = _MASTER_HEADER_MAP.get(norm)
        if canonical:
            matched.add(canonical)
    return matched


def _find_master_header_row(raw_df: pd.DataFrame) -> int:
    """시트 상단에 빈 행이 있어도 헤더 행(상품코드·상품명 포함)을 찾는다."""
    if raw_df.empty:
        return -1
    best_row = -1
    best_score = 0
    max_rows = min(40, len(raw_df))
    for row_idx in range(max_rows):
        matched = _master_header_matches_in_row(raw_df.iloc[row_idx].tolist())
        if "kr_sku" not in matched or "kr_name" not in matched:
            continue
        score = len(matched)
        if score > best_score:
            best_score = score
            best_row = row_idx
    return best_row


def _dataframe_from_master_header_row(raw_df: pd.DataFrame, header_row: int) -> pd.DataFrame:
    """헤더 행 이후를 본문으로 쓰고, 왼쪽 빈 열은 제거한다."""
    ncols = raw_df.shape[1]
    headers: list[str] = []
    for i, val in enumerate(raw_df.iloc[header_row, :ncols].tolist()):
        text = _cell_text(val)
        headers.append(text if text else f"__empty_{i}")

    body = raw_df.iloc[header_row + 1 :, :ncols].copy()
    body.columns = headers
    body = body.reset_index(drop=True)

    keep: list[str] = []
    for col in body.columns:
        if str(col).startswith("__empty_"):
            series = body[col]
            if series.map(lambda v: not _cell_text(v)).all():
                continue
        keep.append(col)
    if keep:
        body = body[keep]
    return body


def _read_master_dataframe(raw: bytes, filename: str) -> tuple[pd.DataFrame, int]:
    """엑셀 → (정규화된 DataFrame, 헤더 행 0-based 인덱스)."""
    last_header_error: HTTPException | None = None
    try:
        raw_df = pd.read_excel(io.BytesIO(raw), sheet_name=0, dtype=object, header=None)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"{filename}: 엑셀을 읽을 수 없습니다. ({exc})") from exc

    header_row = _find_master_header_row(raw_df)
    if header_row >= 0:
        try:
            df = _dataframe_from_master_header_row(raw_df, header_row)
            return _coerce_master_dataframe(df, filename), header_row
        except HTTPException as exc:
            last_header_error = exc

    for fallback_header in (0, 1):
        if fallback_header >= len(raw_df):
            continue
        try:
            df = _dataframe_from_master_header_row(raw_df, fallback_header)
            return _coerce_master_dataframe(df, filename), fallback_header
        except HTTPException as exc:
            if exc.status_code == 400:
                last_header_error = exc
                continue
            raise

    if last_header_error is not None:
        raise last_header_error
    raise HTTPException(
        status_code=400,
        detail=f"{filename}: 인식 가능한 헤더가 없습니다. 필수 열: 상품코드, 상품명, 브랜드",
    )


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
        df, header_row = _read_master_dataframe(raw, filename)
        for idx, row in df.iterrows():
            row_label = f"{filename} {header_row + int(idx) + 2}행"
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
                    "representative_code": _resolve_representative_code(
                        row.get("representative_code") if "representative_code" in df.columns else "",
                        kr_sku,
                    ),
                    "segment": normalize_item_segment(row.get("segment")) if "segment" in df.columns else None,
                    "version": _cell_text(row.get("version"))[:64] if "version" in df.columns else "",
                    "kr_sku": kr_sku,
                    "kr_name": kr_name,
                    "stock_category": _cell_text(row.get("stock_category"))[:64]
                    if "stock_category" in df.columns
                    else "",
                    "fcst_grade": _cell_text(row.get("fcst_grade"))[:32] if "fcst_grade" in df.columns else "",
                    "stock_grade": _cell_text(row.get("stock_grade"))[:32] if "stock_grade" in df.columns else "",
                    "release_month": _normalize_release_month(row.get("release_month"))
                    if "release_month" in df.columns
                    else None,
                    "code_registered_at": _normalize_code_registered_at(row.get("code_registered_at"))
                    if "code_registered_at" in df.columns
                    else None,
                    "kr_grade": _cell_text(row.get("kr_grade"))[:32] if "kr_grade" in df.columns else "",
                    "us_grade": _cell_text(row.get("us_grade"))[:32] if "us_grade" in df.columns else "",
                    "tw_grade": _cell_text(row.get("tw_grade"))[:32] if "tw_grade" in df.columns else "",
                    "hk_grade": _cell_text(row.get("hk_grade"))[:32] if "hk_grade" in df.columns else "",
                    "jp_grade": _cell_text(row.get("jp_grade"))[:32] if "jp_grade" in df.columns else "",
                }
            )
    if not records:
        raise HTTPException(status_code=400, detail="업로드 파일에서 처리할 데이터 행이 없습니다.")
    return records


def _load_kr_locale_pairs(db: Session) -> list[tuple[ProductLocale, ProductGroup]]:
    return list(
        db.execute(
            select(ProductLocale, ProductGroup)
            .join(ProductGroup, ProductLocale.item_id == ProductGroup.id)
            .where(ProductLocale.country_code == "KR")
        ).all()
    )


def _locale_is_active(loc: ProductLocale) -> bool:
    return not sa_inspect(loc).deleted


def _prune_deleted_kr_pairs(pairs: list[tuple[ProductLocale, ProductGroup]]) -> None:
    pairs[:] = [(loc, group) for loc, group in pairs if _locale_is_active(loc)]


def _build_kr_sku_group_map(pairs: list[tuple[ProductLocale, ProductGroup]]) -> dict[str, ProductGroup]:
    """정규화 SKU → ProductGroup (업로드 시 행마다 전체 조회하지 않도록 1회 인덱스)."""
    out: dict[str, ProductGroup] = {}
    for loc, group in pairs:
        if not _locale_is_active(loc):
            continue
        norm = _normalize_mapping_sku(loc.sku)
        if norm and norm not in out:
            out[norm] = group
    return out


def _find_existing_for_master_upload(
    pairs: list[tuple[ProductLocale, ProductGroup]],
    kr_sku_norm: str,
) -> tuple[ProductGroup | None, ProductLocale | None]:
    """업로드 상품코드(정규화)와 일치하는 기존 행 — raw 일치 우선, 없으면 정규화 일치."""
    normalized_hit: tuple[ProductGroup, ProductLocale] | None = None
    for loc, group in pairs:
        if not _locale_is_active(loc):
            continue
        if loc.sku == kr_sku_norm:
            return group, loc
        if normalized_hit is None and _normalize_mapping_sku(loc.sku) == kr_sku_norm:
            normalized_hit = (group, loc)
    if normalized_hit is not None:
        return normalized_hit
    return None, None


def _clear_conflicting_kr_duplicates(
    db: Session,
    pairs: list[tuple[ProductLocale, ProductGroup]],
    sku_group_map: dict[str, ProductGroup],
    *,
    owner_group_id: uuid.UUID,
    kr_sku_norm: str,
) -> None:
    """동일 상품코드(정규화)가 다른 상품에 중복 등록된 KR locale 제거 — 갱신 대상만 남김."""
    remove_at: list[int] = []
    for idx, (loc, group) in enumerate(pairs):
        if group.id == owner_group_id:
            continue
        if loc.sku != kr_sku_norm and _normalize_mapping_sku(loc.sku) != kr_sku_norm:
            continue
        db.delete(loc)
        if loc in group.locales:
            group.locales.remove(loc)
        remove_at.append(idx)
        for key, mapped in list(sku_group_map.items()):
            if mapped.id == group.id:
                del sku_group_map[key]
    for idx in reversed(remove_at):
        pairs.pop(idx)
    _prune_deleted_kr_pairs(pairs)


def _format_master_sku_conflict(
    row_label: str,
    upload_sku: str,
    *,
    upload_name: str = "",
    conflict_loc: ProductLocale,
    conflict_group: ProductGroup,
    reason: str,
) -> str:
    existing_name = (conflict_group.kr_name or conflict_loc.name or "").strip() or "(이름 없음)"
    existing_brand = (conflict_group.brand or "").strip() or "-"
    lines = [
        f"{row_label}: {reason}",
        f"  · 업로드: 코드 {upload_sku!r}" + (f", 상품명 {upload_name!r}" if upload_name else ""),
        f"  · DB 기존: 코드 {conflict_loc.sku!r}, 상품명 {existing_name!r}, 브랜드 {existing_brand!r}",
    ]
    if conflict_loc.sku != upload_sku:
        lines.append(f"  · 참고: 표기는 다르지만 동일 코드로 처리됩니다 ({conflict_loc.sku!r} ↔ {upload_sku!r})")
    return "\n".join(lines)


def _find_kr_sku_conflicts(
    pairs: list[tuple[ProductLocale, ProductGroup]],
    new_sku: str,
    *,
    owner_group_id: uuid.UUID | None,
) -> list[tuple[ProductLocale, ProductGroup]]:
    new_norm = _normalize_mapping_sku(new_sku)
    hits: list[tuple[ProductLocale, ProductGroup]] = []
    for loc, group in pairs:
        if not _locale_is_active(loc):
            continue
        if owner_group_id is not None and group.id == owner_group_id:
            continue
        if loc.sku == new_sku or _normalize_mapping_sku(loc.sku) == new_norm:
            hits.append((loc, group))
    return hits


def _master_upload_integrity_detail(
    exc: IntegrityError,
    *,
    row_label: str,
    kr_sku: str,
    kr_name: str,
    kr_pairs: list[tuple[ProductLocale, ProductGroup]],
) -> str:
    raw = str(getattr(exc, "orig", None) or exc)
    sku_match = re.search(r"\(country_code,\s*sku\)=\([^,]+,\s*([^)]+)\)", raw, re.IGNORECASE)
    conflict_sku = (sku_match.group(1).strip() if sku_match else kr_sku).strip("'\"")
    conflicts = _find_kr_sku_conflicts(kr_pairs, conflict_sku, owner_group_id=None)

    lines = [
        f"{row_label}: 상품코드 {kr_sku!r} 처리 중 DB 중복 오류가 발생했습니다.",
    ]
    if kr_name:
        lines.append(f"  · 업로드 상품명: {kr_name!r}")
    if conflict_sku != kr_sku:
        lines.append(
            f"  · DB 제약 위반 코드: {conflict_sku!r} "
            f"(현재 행과 다르면 같은 파일의 다른 행·열 매핑을 확인하세요)"
        )
    if conflicts:
        for loc, group in conflicts[:5]:
            existing_name = (group.kr_name or loc.name or "").strip() or "(이름 없음)"
            lines.append(
                f"  · 충돌 DB 행: 코드 {loc.sku!r}, 상품명 {existing_name!r}, 브랜드 {(group.brand or '-')!r}"
            )
        if len(conflicts) > 5:
            lines.append(f"  · … 외 {len(conflicts) - 5}건")
    else:
        lines.append(f"  · DB에 이미 한국(KR) 상품코드 {conflict_sku!r} 가 등록되어 있습니다.")
    lines.append(
        f"  · 상품코드 열은 「{'」「'.join(KR_SKU_HEADER_LABELS)}」만 사용합니다 (최종코드 등은 무시)."
    )
    return "\n".join(lines)


def _assert_kr_sku_not_taken_by_other_product(
    pairs: list[tuple[ProductLocale, ProductGroup]],
    *,
    owner_group_id: uuid.UUID | None,
    new_sku: str,
    row_label: str,
    upload_name: str = "",
) -> None:
    """서로 다른 상품이 동일 상품코드를 쓰는 경우만 거부 (갱신·중복 정리 후에도 남은 충돌)."""
    conflicts = _find_kr_sku_conflicts(pairs, new_sku, owner_group_id=owner_group_id)
    if not conflicts:
        return
    loc, group = conflicts[0]
    raise HTTPException(
        status_code=400,
        detail=_format_master_sku_conflict(
            row_label,
            new_sku,
            upload_name=upload_name,
            conflict_loc=loc,
            conflict_group=group,
            reason="다른 상품이 이미 같은 한국 상품코드를 사용 중입니다",
        ),
    )


def _upsert_kr_pair(
    pairs: list[tuple[ProductLocale, ProductGroup]],
    kr_locale: ProductLocale,
    group: ProductGroup,
) -> None:
    for idx, (loc, g) in enumerate(pairs):
        if g.id == group.id:
            pairs[idx] = (kr_locale, group)
            return
    pairs.append((kr_locale, group))


def _find_group_by_kr_sku(db: Session, kr_sku_norm: str) -> ProductGroup | None:
    if not kr_sku_norm:
        return None
    for loc, group in _load_kr_locale_pairs(db):
        if _normalize_mapping_sku(loc.sku) == kr_sku_norm:
            return group
    return None


def _apply_master_fields(group: ProductGroup, kr_locale: ProductLocale, record: dict, now: datetime) -> None:
    group.kr_name = record["kr_name"]
    group.brand = record["brand"]
    group.representative_code = record.get("representative_code") or record["kr_sku"]
    group.segment = record.get("segment")
    group.version = record.get("version") or None
    group.stock_category = record.get("stock_category") or None
    group.fcst_grade = record.get("fcst_grade") or None
    group.stock_grade = record.get("stock_grade") or None
    group.release_month = record.get("release_month") or None
    group.code_registered_at = record.get("code_registered_at")
    group.kr_grade = record.get("kr_grade") or None
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
    kr_pairs = _load_kr_locale_pairs(db)
    sku_group_map = _build_kr_sku_group_map(kr_pairs)
    last_row_label = ""
    last_kr_sku = ""
    last_kr_name = ""
    try:
        for record in records:
            last_row_label = record["row_label"]
            last_kr_sku = record["kr_sku"]
            last_kr_name = record["kr_name"]
            _prune_deleted_kr_pairs(kr_pairs)
            kr_sku = record["kr_sku"]
            existing, existing_kr = _find_existing_for_master_upload(kr_pairs, kr_sku)
            if existing is not None:
                _clear_conflicting_kr_duplicates(
                    db,
                    kr_pairs,
                    sku_group_map,
                    owner_group_id=existing.id,
                    kr_sku_norm=kr_sku,
                )
                _assert_kr_sku_not_taken_by_other_product(
                    kr_pairs,
                    owner_group_id=existing.id,
                    new_sku=kr_sku,
                    row_label=record["row_label"],
                    upload_name=record["kr_name"],
                )
                kr = existing_kr or _kr_locale(existing)
                if kr is None:
                    kr = ProductLocale(
                        id=uuid.uuid4(),
                        item_id=existing.id,
                        country_code="KR",
                        name=record["kr_name"],
                        sku=kr_sku,
                        updated_at=now,
                    )
                    existing.locales.append(kr)
                    db.add(kr)
                _apply_master_fields(existing, kr, record, now)
                _upsert_kr_pair(kr_pairs, kr, existing)
                sku_group_map[kr_sku] = existing
                touched += 1
                _prune_deleted_kr_pairs(kr_pairs)
                continue

            _assert_kr_sku_not_taken_by_other_product(
                kr_pairs,
                owner_group_id=None,
                new_sku=kr_sku,
                row_label=record["row_label"],
                upload_name=record["kr_name"],
            )
            group = ProductGroup(
                id=uuid.uuid4(),
                kr_name=record["kr_name"],
                brand=record["brand"],
                representative_code=record.get("representative_code") or record["kr_sku"],
                segment=record.get("segment"),
                version=record.get("version") or None,
                stock_category=record.get("stock_category") or None,
                fcst_grade=record.get("fcst_grade") or None,
                stock_grade=record.get("stock_grade") or None,
                release_month=record.get("release_month"),
                code_registered_at=record.get("code_registered_at"),
                kr_grade=record.get("kr_grade") or None,
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
            sku_group_map[kr_sku] = group
            kr_pairs.append((kr, group))
            created += 1
            _prune_deleted_kr_pairs(kr_pairs)
        db.commit()
    except HTTPException:
        db.rollback()
        raise
    except IntegrityError as exc:
        db.rollback()
        err = str(exc).lower()
        if "ix_item_mapping_country_sku" in err or "unique" in err:
            raise HTTPException(
                status_code=400,
                detail=_master_upload_integrity_detail(
                    exc,
                    row_label=last_row_label or "(알 수 없는 행)",
                    kr_sku=last_kr_sku,
                    kr_name=last_kr_name,
                    kr_pairs=kr_pairs,
                ),
            ) from exc
        raise
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
        "representative_code": _resolve_representative_code(
            payload.get("representative_code"),
            kr_sku_in,
        ),
        "segment": normalize_item_segment(payload.get("segment"))
        if str(payload.get("segment") or "").strip()
        else None,
        "version": str(payload.get("version") or "").strip()[:64],
        "kr_sku": kr_sku_in,
        "kr_name": kr_name_in,
        "stock_category": str(payload.get("stock_category") or "").strip()[:64],
        "fcst_grade": str(payload.get("fcst_grade") or "").strip()[:32],
        "stock_grade": str(payload.get("stock_grade") or "").strip()[:32],
        "release_month": _normalize_release_month(release_month_in) if release_month_in else None,
        "code_registered_at": _normalize_code_registered_at(payload.get("code_registered_at"))
        if str(payload.get("code_registered_at") or "").strip()
        else None,
        "kr_grade": str(payload.get("kr_grade") or "").strip()[:32],
        "us_grade": str(payload.get("us_grade") or "").strip()[:32],
        "tw_grade": str(payload.get("tw_grade") or "").strip()[:32],
        "hk_grade": str(payload.get("hk_grade") or "").strip()[:32],
        "jp_grade": str(payload.get("jp_grade") or "").strip()[:32],
    }
    now = datetime.now(timezone.utc)
    _apply_master_fields(group, kr, record, now)
    if "barcode" in payload:
        bc = _normalize_barcode_cell(payload.get("barcode"))
        group.barcode = bc if bc else None
    group.manual_updated_at = now

    extra_cols = list_item_master_extra_columns(db)
    allowed_keys = {c["field_key"] for c in extra_cols}
    if "extra_fields" in payload and allowed_keys:
        incoming = payload.get("extra_fields")
        if incoming is not None and not isinstance(incoming, dict):
            raise HTTPException(status_code=400, detail="extra_fields 형식이 올바르지 않습니다.")
        merged = dict(_extra_fields_dict(group.extra_fields))
        if isinstance(incoming, dict):
            for key in allowed_keys:
                if key not in incoming:
                    continue
                val = incoming.get(key)
                if val is None or str(val).strip() == "":
                    merged.pop(key, None)
                else:
                    merged[key] = _cell_text(val)[:EXTRA_FIELD_VALUE_MAX_LEN]
        group.extra_fields = merged or None

    try:
        db.commit()
        fresh = db.execute(
            select(ProductGroup).where(ProductGroup.id == gid).options(selectinload(ProductGroup.locales))
        ).scalar_one()
    except Exception:
        db.rollback()
        raise
    return master_row_payload(fresh, [c["field_key"] for c in extra_cols])


__all__ = [
    "MASTER_COLUMN_LABELS",
    "MASTER_COLUMN_ORDER",
    "get_item_master_row_by_group_id",
    "list_item_master_rows",
    "list_item_master_extra_columns",
    "add_item_master_extra_column",
    "delete_item_master_extra_column",
    "merge_item_master_uploads",
    "master_row_payload",
    "patch_item_master_row",
]
