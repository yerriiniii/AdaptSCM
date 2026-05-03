"""출고 현황: 원시 롱 데이터 → 월·일 집계 → DB(`ShipmentMatrixRow`) → 화면용 피벗.

**업로드**(`persist_shipment_matrix_uploads`): 엑셀은 판매처·상품코드(또는 어드민코드)·주문일·주문수량
4열만 보며 시트/파일명은 무관. 상품코드는 DB 한국 SKU와 매핑·검증하고, 주문일로 월·일 키 합산.
판매처는 `shipment_vendor_channel_maps` 로 1차 구분→상세 칩(자사몰 현황 등) 및 통합 칩(B2C/B2B/해외)에 반영.
"""

from __future__ import annotations

import io
import re
import unicodedata
from collections import defaultdict
from datetime import date, datetime
from typing import Any

import pandas as pd
from fastapi import HTTPException
from sqlalchemy import select
from sqlalchemy.orm import Session

from domains.inventory.models import ShipmentMatrixRow
from domains.inventory.services.inventory_aggregate_service import _normalize_item_code
from domains.inventory.services.shipment_aggregate_service import (
    _load_kr_sku_brand_name,
    normalize_shipment_excel_sku,
)
from domains.inventory.services.shipment_vendor_channel_maps import (
    PRIMARY_TO_DETAIL_CHANNEL,
    PRIMARY_TO_INTEGRATION_CHANNELS,
    normalize_shipment_vendor_key,
    resolve_vendor_primary_with_overrides,
)

# UI 칩 순서와 동일(엑셀 시트 이름과 정확히 일치해야 읽음)
SHIPMENT_MATRIX_SHEET_NAMES: tuple[str, ...] = (
    "B2B 통합",
    "B2C 통합",
    "해외 통합",
    "B2B 현황",
    "쿠팡 현황",
    "올리브영 현황",
    "자사몰 현황",
    "외부몰 현황",
    "대만 현황",
    "홍콩 현황",
    "일본 현황",
    "그 외 해외 현황",
)

DEFAULT_SHIPMENT_MATRIX_YEAR = 2026
_DEFAULT_DATA_YEAR = DEFAULT_SHIPMENT_MATRIX_YEAR

_WEEKDAYS_KO = ("월", "화", "수", "목", "금", "토", "일")


def list_shipment_matrix_years(db: Session) -> list[int]:
    stmt = select(ShipmentMatrixRow.data_year).distinct()
    found = [int(x) for x in db.execute(stmt).scalars().all() if x is not None]
    return sorted(set(found))


def _norm_header_cell(raw: object) -> str:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = unicodedata.normalize("NFKC", str(raw).strip())
    return re.sub(r"\s+", "", s).casefold()


# 원시 출고 첫 시트: 주문 수량 열 — 「주문 수」「주문수량」 등 정규화 후 값
_RAW_ORDER_QTY_HEADER_NORMS: frozenset[str] = frozenset({"주문수량", "주문수"})
# 원시 4열 SKU: 어드민 열과 상품코드 열이 동시에 있으면 어드민만 사용
_RAW_ORDER_SKU_ADMIN_HEADER_NORMS: frozenset[str] = frozenset({"어드민코드"})
_RAW_ORDER_SKU_PRIMARY_HEADER_NORMS: frozenset[str] = frozenset({"상품코드"})
_RAW_ORDER_SKU_HEADER_NORMS: frozenset[str] = frozenset(
    {"상품코드", "어드민코드"}
)  # 시트 판별용(둘 중 하나 있으면 됨)


def _display_dash(val: str | None) -> str:
    s = (val or "").strip()
    return s if s else "–"


def _ko_weekday_label(d: date) -> str:
    return _WEEKDAYS_KO[d.weekday()]


def format_matrix_day_header(d: date) -> str:
    return f"{d.month}/{d.day}({_ko_weekday_label(d)})"


def _parse_int_qty(val: object) -> int | None:
    if val is None:
        return None
    if isinstance(val, float) and pd.isna(val):
        return None
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if isinstance(val, float) and pd.isna(val):
            return None
        try:
            return int(round(float(val)))
        except (TypeError, ValueError):
            return None
    s = str(val).strip()
    if not s or s in {"-", "—", "–"}:
        return None
    s = s.replace(",", "").replace(" ", "")
    if re.fullmatch(r"-?\d+", s):
        return int(s)
    try:
        return int(round(float(s)))
    except (TypeError, ValueError):
        return None


def _merge_qty_maps(base: dict[str, int], incoming: dict[str, int], *, add: bool = False) -> None:
    for k, v in incoming.items():
        if v is None:
            continue
        if add:
            base[k] = int(base.get(k, 0)) + int(v)
        else:
            base[k] = int(v)


def _first_raw_orders_sheet_name(xl: pd.ExcelFile) -> str | None:
    """파일명·시트 순서 무관: 필수 4열 헤더가 있는 첫 시트."""
    for name in xl.sheet_names or []:
        try:
            df0 = pd.read_excel(xl, sheet_name=name, header=0, nrows=0, dtype=object)
        except Exception:
            continue
        cols = {_norm_header_cell(c) for c in df0.columns}
        if not {"판매처", "주문일"}.issubset(cols):
            continue
        if not (cols & _RAW_ORDER_SKU_HEADER_NORMS):
            continue
        if cols & _RAW_ORDER_QTY_HEADER_NORMS:
            return name
    return None


def shipment_workbook_is_raw_orders_format(xl: pd.ExcelFile) -> bool:
    """원시 4열: 판매처·상품코드·주문일·주문수량(또는 주문 수). 어느 시트든 헤더만 맞으면 됨."""
    return _first_raw_orders_sheet_name(xl) is not None


def _excel_data_row_1based(pos_zero_based: int) -> int:
    """read_excel(header=0) 기준: pos 0 → 엑셀 2행(헤더 다음 첫 데이터 행)."""
    return int(pos_zero_based) + 2


def _is_shipment_total_summary_vendor_cell(raw: object) -> bool:
    """요약 행: 판매처 칸이 「합계」(공백·전각만 다른 표기 포함)이면 집계하지 않는다."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return False
    s = unicodedata.normalize("NFKC", str(raw).strip())
    s = re.sub(r"\s+", "", s)
    return s == "합계"


def parse_shipment_raw_orders_workbook(
    db: Session,
    raw: bytes,
    filename: str = "",
    *,
    kr_sku_brand_name: dict[str, tuple[str, str, str, str]] | None = None,
    vendor_primary_overrides_internal: dict[str, str] | None = None,
) -> tuple[dict[str, list[dict[str, Any]]], int]:
    """판매처·상품코드·주문일·주문수량 롱 데이터 → 칩(시트)별 매트릭스 레코드 + 데이터 연도.

    동일 상품·칩에 대해 월·일 키별 수량을 합산한다.
    """
    bio = io.BytesIO(raw)
    xl = pd.ExcelFile(bio)
    sheet_name = _first_raw_orders_sheet_name(xl)
    if sheet_name is None:
        raise HTTPException(
            status_code=400,
            detail={
                "code": "SHIPMENT_WRONG_FORMAT",
                "message": (
                    f"{filename or '파일'}: 원시 출고 형식이 아닙니다. "
                    "시트 중 하나에 「판매처」「상품코드」(또는 「어드민코드」/「어드민 코드」)「주문일」「주문수량」(또는 「주문 수」) 열이 있어야 합니다."
                ),
                "filename": (filename or "").strip(),
                "sheet_name": None,
            },
        )

    df = pd.read_excel(xl, sheet_name=sheet_name, header=0, dtype=object)
    col_vendor = col_date = col_qty = None
    col_sku_admin = col_sku_primary = None
    for c in df.columns:
        n = _norm_header_cell(c)
        if n == "판매처":
            col_vendor = c
        elif n in _RAW_ORDER_SKU_ADMIN_HEADER_NORMS:
            col_sku_admin = c
        elif n in _RAW_ORDER_SKU_PRIMARY_HEADER_NORMS:
            col_sku_primary = c
        elif n == "주문일":
            col_date = c
        elif n in _RAW_ORDER_QTY_HEADER_NORMS:
            col_qty = c
    col_sku = col_sku_admin if col_sku_admin is not None else col_sku_primary
    if col_vendor is None or col_sku is None or col_date is None or col_qty is None:
        raise HTTPException(
            status_code=400,
            detail={
                "code": "SHIPMENT_MISSING_COLUMNS",
                "message": f"{filename or '파일'}: 시트 「{sheet_name}」에서 필수 열을 찾지 못했습니다.",
                "filename": (filename or "").strip(),
                "sheet_name": sheet_name,
            },
        )

    sku_map = kr_sku_brand_name if kr_sku_brand_name is not None else _load_kr_sku_brand_name(db)
    missing_sku_rows: dict[str, set[int]] = defaultdict(set)
    unknown_vendor_by_norm: dict[str, dict[str, Any]] = {}

    # channel_sheet -> sku -> { monthly: {str month}, daily: {iso} }
    acc: dict[str, dict[str, dict[str, dict[str, int]]]] = defaultdict(
        lambda: defaultdict(lambda: {"monthly": {}, "daily": {}})
    )
    years_seen: list[int] = []

    for pos, (_, row) in enumerate(df.iterrows()):
        excel_row = _excel_data_row_1based(pos)
        vendor = row.get(col_vendor)
        if vendor is None or (isinstance(vendor, float) and pd.isna(vendor)):
            continue
        if _is_shipment_total_summary_vendor_cell(vendor):
            continue
        primary = resolve_vendor_primary_with_overrides(vendor, vendor_primary_overrides_internal)
        if primary is None:
            nk = normalize_shipment_vendor_key(vendor)
            if nk:
                if nk not in unknown_vendor_by_norm:
                    unknown_vendor_by_norm[nk] = {"display": str(vendor).strip(), "rows": set()}
                unknown_vendor_by_norm[nk]["rows"].add(excel_row)
            continue
        if primary == "제외":
            continue

        detail_ch = PRIMARY_TO_DETAIL_CHANNEL.get(primary)
        if not detail_ch:
            continue
        extra_ch = PRIMARY_TO_INTEGRATION_CHANNELS.get(primary, ())
        target_channels = (detail_ch, *extra_ch)

        ex_norm = normalize_shipment_excel_sku(row.get(col_sku))
        if not ex_norm:
            continue
        canon = _normalize_item_code(ex_norm)
        if not canon or canon not in sku_map:
            missing_sku_rows[str(ex_norm).strip()].add(excel_row)
            continue

        raw_dt = row.get(col_date)
        if raw_dt is None or (isinstance(raw_dt, float) and pd.isna(raw_dt)):
            continue
        if isinstance(raw_dt, datetime):
            d_only = raw_dt.date()
        elif isinstance(raw_dt, date) and not isinstance(raw_dt, datetime):
            d_only = raw_dt
        else:
            try:
                parsed = pd.to_datetime(raw_dt, errors="coerce")
                if pd.isna(parsed):
                    continue
                d_only = pd.Timestamp(parsed).date()
            except Exception:
                continue

        years_seen.append(d_only.year)
        mo_key = str(d_only.month)
        day_key = d_only.isoformat()

        q = _parse_int_qty(row.get(col_qty))
        if q is None or q == 0:
            continue

        for ch in target_channels:
            if not ch:
                continue
            bucket = acc[ch][canon]
            _merge_qty_maps(bucket["monthly"], {mo_key: q}, add=True)
            _merge_qty_maps(bucket["daily"], {day_key: q}, add=True)

    if unknown_vendor_by_norm:
        vendor_issues: list[dict[str, Any]] = []
        vendors_list: list[str] = []
        for nk in sorted(unknown_vendor_by_norm.keys()):
            ent = unknown_vendor_by_norm[nk]
            rows_sorted = sorted(ent["rows"])
            disp = str(ent["display"] or "").strip()
            vendor_issues.append({"vendor": disp, "excel_rows": rows_sorted})
            vendors_list.append(disp)
        raise HTTPException(
            status_code=400,
            detail={
                "code": "UNKNOWN_SHIPMENT_VENDORS",
                "message": "표에 없는 판매처가 있습니다. 각 판매처의 구분을 선택한 뒤 다시 시도해 주세요.",
                "filename": (filename or "").strip(),
                "sheet_name": sheet_name,
                "vendors": vendors_list,
                "vendor_issues": vendor_issues,
            },
        )

    if missing_sku_rows:
        sku_issues = [
            {"sku": sku, "excel_rows": sorted(missing_sku_rows[sku])}
            for sku in sorted(missing_sku_rows.keys())
        ]
        raise HTTPException(
            status_code=400,
            detail={
                "code": "UNKNOWN_SKUS",
                "message": (
                    "데이터베이스에 등록되어 있지 않은 상품코드가 있습니다. "
                    "아래 코드를 SKU 관리 탭에서 한국 SKU로 먼저 등록한 뒤 다시 업로드해 주세요."
                ),
                "filename": (filename or "").strip(),
                "sheet_name": sheet_name,
                "skus": sorted(missing_sku_rows.keys()),
                "sku_issues": sku_issues,
            },
        )

    inferred_year = max(years_seen) if years_seen else _DEFAULT_DATA_YEAR
    if inferred_year < 2000 or inferred_year > 2100:
        inferred_year = _DEFAULT_DATA_YEAR

    out: dict[str, list[dict[str, Any]]] = {}
    for ch_name, by_sku in acc.items():
        rows_ch: list[dict[str, Any]] = []
        for sku_key, packs in by_sku.items():
            rows_ch.append(
                {
                    "sku": sku_key,
                    "mkt_priority": "",
                    "category": "",
                    "monthly": {str(k): int(v) for k, v in packs["monthly"].items()},
                    "daily": {str(k): int(v) for k, v in packs["daily"].items()},
                }
            )
        if rows_ch:
            out[ch_name] = rows_ch

    if not out:
        raise HTTPException(
            status_code=400,
            detail={
                "code": "SHIPMENT_NO_AGGREGATABLE_ROWS",
                "message": (
                    f"{filename or '파일'}: 시트 「{sheet_name}」에서 집계할 출고 행이 없습니다. "
                    "(제외 판매처·수량 0·날짜 누락 등)"
                ),
                "filename": (filename or "").strip(),
                "sheet_name": sheet_name,
            },
        )

    return out, inferred_year


def _distinct_channels_with_data(db: Session, data_year: int) -> list[str]:
    stmt = (
        select(ShipmentMatrixRow.channel_sheet)
        .where(ShipmentMatrixRow.data_year == data_year)
        .distinct()
    )
    found = {str(x) for x in db.execute(stmt).scalars().all() if x}
    return [name for name in SHIPMENT_MATRIX_SHEET_NAMES if name in found]


def build_shipment_matrix_view(
    db: Session,
    *,
    channel: str | None = None,
    data_year: int = _DEFAULT_DATA_YEAR,
    all_channels: bool = False,
) -> dict[str, Any]:
    channels_with = _distinct_channels_with_data(db, data_year)
    chip_order = list(SHIPMENT_MATRIX_SHEET_NAMES)
    ch_rank = {name: i for i, name in enumerate(chip_order)}
    active = channel if channel in channels_with else (channels_with[0] if channels_with else None)

    if not channels_with:
        return {
            "summary": {"item_count": 0, "date_count": 0},
            "countries": ["KR"],
            "dates": [],
            "rows": [],
            "channels": chip_order,
            "shipment_month_columns": [],
            "shipment_day_labels": {},
            "shipment_totals": {"monthly": {}, "daily": {}},
            "shipment_active_channel": None,
            "shipment_channels_with_data": [],
        }

    if all_channels:
        rows_orm = (
            db.execute(
                select(ShipmentMatrixRow).where(
                    ShipmentMatrixRow.data_year == data_year,
                )
            )
            .scalars()
            .all()
        )
        matrix_active: str | None = None
    else:
        if not active:
            return {
                "summary": {"item_count": 0, "date_count": 0},
                "countries": ["KR"],
                "dates": [],
                "rows": [],
                "channels": chip_order,
                "shipment_month_columns": [],
                "shipment_day_labels": {},
                "shipment_totals": {"monthly": {}, "daily": {}},
                "shipment_active_channel": None,
                "shipment_channels_with_data": [],
            }
        rows_orm = (
            db.execute(
                select(ShipmentMatrixRow).where(
                    ShipmentMatrixRow.channel_sheet == active,
                    ShipmentMatrixRow.data_year == data_year,
                )
            )
            .scalars()
            .all()
        )
        matrix_active = active

    month_keys: set[str] = set()
    day_keys: set[str] = set()
    for rec in rows_orm:
        for k in (rec.monthly_totals or {}):
            month_keys.add(str(k))
        for k in (rec.daily_totals or {}):
            day_keys.add(str(k))

    month_ints = sorted({int(m) for m in month_keys if m.isdigit()})
    shipment_month_columns = [{"key": str(m), "label": f"{m}월 합계"} for m in month_ints]

    day_sorted = sorted(day_keys)
    shipment_day_labels: dict[str, str] = {}
    for iso in day_sorted:
        try:
            y, mo, d = map(int, iso.split("-"))
            shipment_day_labels[iso] = format_matrix_day_header(date(y, mo, d))
        except ValueError:
            shipment_day_labels[iso] = iso

    sku_map = _load_kr_sku_brand_name(db)
    out_rows: list[dict[str, Any]] = []
    sum_monthly: dict[str, int] = defaultdict(int)
    sum_daily: dict[str, int] = defaultdict(int)

    def _rec_channel(rec: ShipmentMatrixRow) -> str:
        return str(rec.channel_sheet or "").strip() or (matrix_active or "")

    for rec in sorted(
        rows_orm,
        key=lambda x: (ch_rank.get(str(x.channel_sheet or ""), 999), x.sku),
    ):
        _b, _n, db_mkt, db_seg = sku_map.get(rec.sku, ("", "", "", ""))
        brand, pname = _b, _n
        mt = {str(k): int(v) for k, v in (rec.monthly_totals or {}).items()}
        dt = {str(k): int(v) for k, v in (rec.daily_totals or {}).items()}
        for k, v in mt.items():
            sum_monthly[k] += v
        for k, v in dt.items():
            sum_daily[k] += v

        row_channel = _rec_channel(rec)
        mkt_show = db_mkt or (str(rec.mkt_priority or "").strip())
        seg_show = db_seg or (str(rec.category or "").strip())
        flat: dict[str, Any] = {
            "sku": rec.sku,
            "brand": brand,
            "description": pname,
            "mkt_priority": _display_dash(mkt_show or None),
            "segment": _display_dash(seg_show or None),
            "country": "KR",
            "channel": row_channel,
            "month_totals": mt,
        }
        for dk, q in dt.items():
            flat[dk] = q
        out_rows.append(flat)

    dates_out = day_sorted
    totals_channel = "" if all_channels else (matrix_active or "")
    totals_row: dict[str, Any] = {
        "is_total": True,
        "sku": "",
        "brand": "",
        "description": "합 계",
        "mkt_priority": "",
        "segment": "",
        "country": "KR",
        "channel": totals_channel,
        "month_totals": {k: sum_monthly.get(k, 0) for k in [str(m) for m in month_ints]},
    }
    for dk in dates_out:
        totals_row[dk] = sum_daily.get(dk, 0)

    return {
        "summary": {"item_count": len(out_rows), "date_count": len(dates_out)},
        "countries": ["KR"],
        "dates": dates_out,
        "rows": out_rows,
        "channels": chip_order,
        "shipment_month_columns": shipment_month_columns,
        "shipment_day_labels": shipment_day_labels,
        "shipment_totals": {"monthly": dict(sum_monthly), "daily": dict(sum_daily)},
        "shipment_active_channel": None if all_channels else matrix_active,
        "shipment_channels_with_data": channels_with,
        "shipment_totals_row": totals_row,
    }
