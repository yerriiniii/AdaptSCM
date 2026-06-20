"""출고 적로업공용: 국한 상품코드(국규및 한국 SKU나 브랜드·표준명) 조회.

판매처 1단 구분·모집계는 `shipment_vendor_channel_maps` 와 `shipment_matrix` 가 담당한다.
"""

from __future__ import annotations

import re

from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from domains.item.models import ProductGroup, ProductLocale
from shared.sku import normalize_item_code


def normalize_shipment_excel_sku(raw: object) -> str:
    s = str(raw or "").strip()
    if not s:
        return ""
    s = re.sub(r"\.0$", "", s)
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    up = s.upper()
    if up.startswith("S"):
        return up
    if re.fullmatch(r"\d+", s):
        if len(s) <= 4:
            return s.zfill(5)
        return s
    return up


def _load_kr_sku_brand_name(db: Session) -> dict[str, tuple[str, str, str, str]]:
    """normalize_item_code(sku) -> (brand, kr_display_name, mkt_priority, segment).

    mkt_priority·segment 는 상품 그룹(`item`) 마스터에서.
    """
    locales = (
        db.execute(
            select(ProductLocale)
            .options(selectinload(ProductLocale.product_group))
            .where(ProductLocale.country_code == "KR")
        )
        .scalars()
        .all()
    )
    out: dict[str, tuple[str, str, str, str]] = {}
    for loc in locales:
        key = normalize_item_code(loc.sku)
        if not key:
            continue
        pg = loc.product_group
        brand = (pg.brand or "").strip() if pg else ""
        name = (loc.name or "").strip()
        if not name and pg:
            name = (pg.kr_name or "").strip()
        mkt = (pg.mkt_priority or "").strip() if pg else ""
        seg = (pg.segment or "").strip() if pg else ""
        out[key] = (brand, name, mkt, seg)
    return out
