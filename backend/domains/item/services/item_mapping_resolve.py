"""재고 행 ↔ item / item_mapping 해석 (국가별 SKU 고정, 업로드 시 미등록 거절)."""

from __future__ import annotations

from typing import Any

import pandas as pd
from fastapi import HTTPException
from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from domains.item.models import ProductGroup, ProductLocale
from shared.sku import normalize_item_code

UNKNOWN_SKUS_DETAIL_CODE = "UNKNOWN_SKUS"


def _normalized_sku_allowlist_for_country(db: Session, country_code: str, cache: dict[str, set[str]]) -> set[str]:
    cc = str(country_code or "").strip().upper()
    if cc not in cache:
        skus = db.execute(select(ProductLocale.sku).where(ProductLocale.country_code == cc)).scalars().all()
        cache[cc] = {normalize_item_code(s) for s in skus if s}
    return cache[cc]


def validate_stock_row_frame_item_mapping(
    db: Session,
    row_frame: pd.DataFrame,
    country_code: str,
    source_label: str,
    *,
    allowlist_cache: dict[str, set[str]],
) -> None:
    """국가별 SKU가 item_mapping 에 있어야 한다. 없으면 400 과 목록 반환."""
    cc = str(country_code or "").strip().upper()
    if row_frame.empty:
        return
    allowed = _normalized_sku_allowlist_for_country(db, cc, allowlist_cache)
    empty_rows: list[str] = []
    missing_pairs: set[tuple[str, str]] = set()
    for record in row_frame.itertuples(index=False):
        raw_sku = str(getattr(record, "sku", "") or "").strip()
        norm = normalize_item_code(raw_sku)
        desc = str(getattr(record, "description", "") or "").strip()
        desc_bit = f" 상품명={desc!r}" if desc else ""
        if not norm:
            empty_rows.append(f"{source_label}: SKU가 비어 있습니다.{desc_bit}")
            continue
        if norm not in allowed:
            missing_pairs.add((cc, norm))
    if missing_pairs:
        items: list[dict[str, str]] = [
            {"country_code": c, "sku": s} for c, s in sorted(missing_pairs)
        ]
        detail: dict[str, Any] = {
            "code": UNKNOWN_SKUS_DETAIL_CODE,
            "message": (
                "데이터베이스에 등록되어 있지 않은 상품코드가 있습니다. "
                "아래 국가·코드를 SKU 관리 탭에서 먼저 등록한 뒤 재고 통합을 다시 실행해 주세요."
            ),
            "items": items,
        }
        if empty_rows:
            detail["sku_empty_rows"] = sorted(empty_rows)
        raise HTTPException(status_code=400, detail=detail)
    if empty_rows:
        raise HTTPException(status_code=400, detail=sorted(empty_rows))


def _locale_display_and_brand_maps_for_country(db: Session, country_code: str) -> tuple[dict[str, str], dict[str, str]]:
    """정규화 SKU -> (해당 국가 ProductLocale 표시명, ProductGroup.brand)."""
    cc = str(country_code or "").strip().upper()
    name_by: dict[str, str] = {}
    brand_by: dict[str, str] = {}
    locales = (
        db.execute(
            select(ProductLocale)
            .options(selectinload(ProductLocale.product_group))
            .where(ProductLocale.country_code == cc)
        )
        .scalars()
        .all()
    )
    for loc in locales:
        n = normalize_item_code(loc.sku)
        if not n:
            continue
        group = loc.product_group
        disp = (loc.name or "").strip()
        if not disp and group:
            disp = (group.kr_name or "").strip()
        name_by[n] = disp
        brand_by[n] = (group.brand or "").strip() if group else ""
    return name_by, brand_by


def apply_db_locale_columns_to_raw_stock_df(db: Session, df: pd.DataFrame, country_code: str) -> None:
    """엑셀 상품명·공급처를 DB 마스터로 덮어쓴다. _read_inventory_file 직후 df(sku·description·supplier)용. inplace."""
    if df.empty or "sku" not in df.columns:
        return
    cc = str(country_code or "").strip().upper()
    name_by, brand_by = _locale_display_and_brand_maps_for_country(db, cc)
    if not name_by:
        return
    norms = df["sku"].map(normalize_item_code)
    desc_excel = df["description"].fillna("").astype(str).str.strip()
    sup_excel = df["supplier"].fillna("").astype(str).str.strip()
    md = norms.map(name_by)
    mb = norms.map(brand_by)
    df["description"] = md.where(md.notna() & md.astype(str).str.strip().ne(""), desc_excel)
    df["supplier"] = mb.where(mb.notna() & mb.astype(str).str.strip().ne(""), sup_excel)


def apply_db_locale_columns_to_stock_merged(db: Session, merged: pd.DataFrame) -> None:
    """집계 직전 long 프레임: 국가별로 상품명·공급처를 DB 기준으로 맞춘 뒤, 동일 키 행이 합쳐지도록 한다. inplace."""
    if merged.empty or "country" not in merged.columns or "sku" not in merged.columns:
        return
    for cc in merged["country"].dropna().unique():
        cc_u = str(cc).upper().strip()
        if not cc_u or cc_u == "SHIPMENT":
            continue
        idx = merged["country"] == cc_u
        if not idx.any():
            continue
        name_by, brand_by = _locale_display_and_brand_maps_for_country(db, cc_u)
        if not name_by:
            continue
        norms = merged.loc[idx, "sku"].map(normalize_item_code)
        desc_excel = merged.loc[idx, "description"].fillna("").astype(str).str.strip()
        sup_excel = merged.loc[idx, "supplier"].fillna("").astype(str).str.strip()
        md = norms.map(name_by)
        mb = norms.map(brand_by)
        merged.loc[idx, "description"] = md.where(md.notna() & md.astype(str).str.strip().ne(""), desc_excel).values
        merged.loc[idx, "supplier"] = mb.where(mb.notna() & mb.astype(str).str.strip().ne(""), sup_excel).values


def _kr_display_for_locale_row(locale_row: ProductLocale) -> tuple[str, str]:
    """(mapped_kr_name, mapped_kr_sku). KR SKU 없으면 이름은 '-'."""
    group = locale_row.product_group
    kr = next((loc for loc in group.locales if loc.country_code == "KR"), None)
    kr_sku_n = normalize_item_code(kr.sku) if kr else ""
    if not kr_sku_n:
        return "-", ""
    name = (kr.name or "").strip() if kr else ""
    display_name = name or (group.kr_name or "").strip() or "-"
    return display_name, kr_sku_n


def attach_item_mapping_kr_to_stock_rows(db: Session, country_code: str, result_rows: list[dict]) -> None:
    """피벗 결과 행에 mapped_kr_name / mapped_kr_sku 를 item_mapping 기준으로 붙인다."""
    cc = str(country_code or "").strip().upper()
    if not result_rows:
        return
    locales = (
        db.execute(
            select(ProductLocale)
            .options(
                selectinload(ProductLocale.product_group).selectinload(ProductGroup.locales),
            )
            .where(ProductLocale.country_code == cc)
        )
        .scalars()
        .all()
    )
    by_nsku: dict[str, ProductLocale] = {}
    for loc in locales:
        key = normalize_item_code(loc.sku)
        if key:
            by_nsku[key] = loc

    for row in result_rows:
        nsku = normalize_item_code(row.get("sku"))
        loc = by_nsku.get(nsku) if nsku else None
        if loc is None:
            row["mapped_kr_name"] = "-"
            row["mapped_kr_sku"] = ""
            row["mapped_barcode"] = ""
            continue
        kr_name, kr_sku = _kr_display_for_locale_row(loc)
        row["mapped_kr_name"] = kr_name
        row["mapped_kr_sku"] = kr_sku
        group = loc.product_group
        row["mapped_barcode"] = str(group.barcode or "").strip() if group else ""


__all__ = [
    "UNKNOWN_SKUS_DETAIL_CODE",
    "apply_db_locale_columns_to_raw_stock_df",
    "apply_db_locale_columns_to_stock_merged",
    "attach_item_mapping_kr_to_stock_rows",
    "validate_stock_row_frame_item_mapping",
]
