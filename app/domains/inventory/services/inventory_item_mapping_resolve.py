"""재고 행 ↔ item / item_mapping 해석 (국가별 SKU 고정, 업로드 시 미등록 거절)."""

from __future__ import annotations

import pandas as pd
from fastapi import HTTPException
from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from app.domains.inventory.models import ProductGroup, ProductLocale
from app.domains.inventory.services.inventory_aggregate_service import _normalize_item_code


def _normalized_sku_allowlist_for_country(db: Session, country_code: str, cache: dict[str, set[str]]) -> set[str]:
    cc = str(country_code or "").strip().upper()
    if cc not in cache:
        skus = db.execute(select(ProductLocale.sku).where(ProductLocale.country_code == cc)).scalars().all()
        cache[cc] = {_normalize_item_code(s) for s in skus if s}
    return cache[cc]


def validate_inventory_row_frame_item_mapping(
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
    unknown: set[str] = set()
    for record in row_frame.itertuples(index=False):
        raw_sku = str(getattr(record, "sku", "") or "").strip()
        norm = _normalize_item_code(raw_sku)
        desc = str(getattr(record, "description", "") or "").strip()
        desc_bit = f" 상품명={desc!r}" if desc else ""
        if not norm:
            unknown.add(f"{source_label}: SKU가 비어 있습니다.{desc_bit}")
            continue
        if norm not in allowed:
            unknown.add(f"{source_label}: 미등록 SKU 국가={cc} SKU={norm!r}{desc_bit}")
    if unknown:
        lines = [
            "item / item_mapping 에 등록되지 않은 상품이 있습니다. SKU 마스터에서 해당 국가 SKU를 먼저 등록한 뒤 재고 파일을 올려주세요.",
            *sorted(unknown),
        ]
        raise HTTPException(status_code=400, detail=lines)


def _kr_display_for_locale_row(locale_row: ProductLocale) -> tuple[str, str]:
    """(mapped_kr_name, mapped_kr_sku). KR SKU 없으면 이름은 '-'."""
    group = locale_row.product_group
    kr = next((loc for loc in group.locales if loc.country_code == "KR"), None)
    kr_sku_n = _normalize_item_code(kr.sku) if kr else ""
    if not kr_sku_n:
        return "-", ""
    name = (kr.name or "").strip() if kr else ""
    display_name = name or (group.kr_name or "").strip() or "-"
    return display_name, kr_sku_n


def attach_item_mapping_kr_to_inventory_rows(db: Session, country_code: str, result_rows: list[dict]) -> None:
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
        key = _normalize_item_code(loc.sku)
        if key:
            by_nsku[key] = loc

    for row in result_rows:
        nsku = _normalize_item_code(row.get("sku"))
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
