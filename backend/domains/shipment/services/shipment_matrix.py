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

from domains.shipment.models import ShipmentMatrixRow
from shared.sku import normalize_item_code
from domains.shipment.services.shipment_aggregate import (
    _load_kr_sku_brand_name,
    normalize_shipment_excel_sku,
)
from domains.shipment.services.shipment_vendor_channel_maps import (
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
# 원시 4열 SKU: 어드민·상품코드 열이 둘 다 있으면 기본은 어드민 셀.
# 행마다 어드민이 비었거나 「미사용」(화살표·공백 표기 포함)이면 그 행만 상품코드 셀.
_RAW_ORDER_SKU_ADMIN_HEADER_NORMS: frozenset[str] = frozenset({"어드민코드"})
_RAW_ORDER_SKU_PRIMARY_HEADER_NORMS: frozenset[str] = frozenset({"상품코드"})
_RAW_ORDER_MKT_HEADER_NORMS: frozenset[str] = frozenset(
    {"마케팅우선순위", "마케팅등급", "mktpriority", "mkt"}
)
_RAW_ORDER_SEGMENT_HEADER_NORMS: frozenset[str] = frozenset(
    {"구분", "segment", "카테고리", "분류"}
)
_RAW_ORDER_SKU_HEADER_NORMS: frozenset[str] = frozenset(
    {"상품코드", "어드민코드"}
)  # 시트 판별용(둘 중 하나 있으면 됨)
_NO_ADMIN_SKU_PREFIX = "__NO_ADMIN__::"


def _shipment_admin_sku_falls_back_to_primary(raw: object) -> bool:
    """어드민코드 셀이 비었거나 '미사용' 등이면 같은 행의 상품코드로 매핑한다."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return True
    s = unicodedata.normalize("NFKC", str(raw).strip())
    if not s or s in {"-", "—", "–"}:
        return True
    compact = re.sub(r"\s+", "", s)
    compact = re.sub(r"^[<\u003c←—–\-・]+", "", compact)
    compact = re.sub(r"[>\u003e→—–\-・]+$", "", compact)
    return compact == "" or compact == "미사용"


def _resolve_shipment_row_sku_cell(
    row: pd.Series,
    col_sku_admin: str | None,
    col_sku_primary: str | None,
) -> object:
    """열 선택: 어드민 우선. 다만 행마다 어드민이 비어 있거나 미사용이면 상품코드 셀."""
    if col_sku_admin is not None:
        admin_raw = row.get(col_sku_admin)
        if not _shipment_admin_sku_falls_back_to_primary(admin_raw):
            return admin_raw
        if col_sku_primary is not None:
            return row.get(col_sku_primary)
        return admin_raw
    if col_sku_primary is not None:
        return row.get(col_sku_primary)
    return None


def _make_no_admin_storage_sku(primary_sku: str) -> str:
    return f"{_NO_ADMIN_SKU_PREFIX}{primary_sku}"


def _parse_no_admin_storage_sku(storage_sku: str) -> str | None:
    s = str(storage_sku or "").strip()
    if not s.startswith(_NO_ADMIN_SKU_PREFIX):
        return None
    out = s[len(_NO_ADMIN_SKU_PREFIX) :].strip()
    return out or None


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


_EXCEL_SERIAL_ORIGIN = pd.Timestamp("1899-12-30")
_SHIPMENT_DATE_YEAR_MIN = 1990
_SHIPMENT_DATE_YEAR_MAX = 2100


def _parse_shipment_order_date(raw: object) -> date | None:
    """주문일 셀 → date. ISO, 한글(년월일), 2자리 연(26-01-01), 엑셀 직렬값 등."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    if isinstance(raw, datetime):
        d = raw.date()
        return d if _SHIPMENT_DATE_YEAR_MIN <= d.year <= _SHIPMENT_DATE_YEAR_MAX else None
    if isinstance(raw, date) and not isinstance(raw, datetime):
        return raw if _SHIPMENT_DATE_YEAR_MIN <= raw.year <= _SHIPMENT_DATE_YEAR_MAX else None

    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        x = float(raw)
        if 1e6 < x < 1e8:
            n = int(round(x))
            s8 = f"{n:08d}"
            try:
                d = date(int(s8[0:4]), int(s8[4:6]), int(s8[6:8]))
                if _SHIPMENT_DATE_YEAR_MIN <= d.year <= _SHIPMENT_DATE_YEAR_MAX:
                    return d
            except ValueError:
                pass
        if 1 <= x < 1e6:
            try:
                ts = _EXCEL_SERIAL_ORIGIN + pd.Timedelta(days=x)
                if pd.isna(ts):
                    return None
                d = pd.Timestamp(ts).date()
                if _SHIPMENT_DATE_YEAR_MIN <= d.year <= _SHIPMENT_DATE_YEAR_MAX:
                    return d
            except (ValueError, OverflowError, OSError):
                pass

    s = unicodedata.normalize("NFKC", str(raw).strip())
    if not s:
        return None

    m_ko4 = re.match(r"^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일?\s*$", s)
    if m_ko4:
        try:
            d = date(int(m_ko4.group(1)), int(m_ko4.group(2)), int(m_ko4.group(3)))
            if _SHIPMENT_DATE_YEAR_MIN <= d.year <= _SHIPMENT_DATE_YEAR_MAX:
                return d
        except ValueError:
            return None

    m_ko2 = re.match(r"^(\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일?\s*$", s)
    if m_ko2:
        yy = int(m_ko2.group(1))
        y = 2000 + yy if yy < 100 else yy
        try:
            d = date(y, int(m_ko2.group(2)), int(m_ko2.group(3)))
            if _SHIPMENT_DATE_YEAR_MIN <= d.year <= _SHIPMENT_DATE_YEAR_MAX:
                return d
        except ValueError:
            return None

    # 4자리 연도(2026-01-01)를 2자리 패턴(26-01-01)보다 먼저 — "2026-01-01" 오인식 방지
    m_iso = re.match(r"^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:\s+|$)", s)
    if m_iso:
        try:
            d = date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3)))
            if _SHIPMENT_DATE_YEAR_MIN <= d.year <= _SHIPMENT_DATE_YEAR_MAX:
                return d
        except ValueError:
            return None

    m_short = re.match(r"^(\d{2})[-/.](\d{1,2})[-/.](\d{1,2})\s*$", s)
    if m_short:
        yy = int(m_short.group(1))
        y = 2000 + yy if yy < 100 else yy
        try:
            d = date(y, int(m_short.group(2)), int(m_short.group(3)))
            if _SHIPMENT_DATE_YEAR_MIN <= d.year <= _SHIPMENT_DATE_YEAR_MAX:
                return d
        except ValueError:
            return None

    try:
        parsed = pd.to_datetime(s, errors="coerce")
        if pd.isna(parsed):
            return None
        d = pd.Timestamp(parsed).date()
        if _SHIPMENT_DATE_YEAR_MIN <= d.year <= _SHIPMENT_DATE_YEAR_MAX:
            return d
    except Exception:
        return None
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
) -> tuple[dict[int, dict[str, list[dict[str, Any]]]], list[int]]:
    """판매처·상품코드·주문일·주문수량 롱 데이터 → 연도별 칩(시트)별 매트릭스 레코드.

    주문일 연도가 여러 개면 연도마다 분리해 반환한다(월 합계·일자 키가 연도별로 섞이지 않음).
    동일 연도·상품·칩에 대해 월·일 키별 수량을 합산한다.
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
    col_mkt = col_segment = None
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
        elif n in _RAW_ORDER_MKT_HEADER_NORMS:
            col_mkt = c
        elif n in _RAW_ORDER_SEGMENT_HEADER_NORMS:
            col_segment = c
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

    # data_year -> channel_sheet -> sku -> { monthly, daily, mkt_priority, category }
    acc: dict[int, dict[str, dict[str, dict[str, Any]]]] = defaultdict(
        lambda: defaultdict(lambda: defaultdict(lambda: {"monthly": {}, "daily": {}, "mkt_priority": "", "category": ""}))
    )

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

        row_uses_admin_fallback = False
        if col_sku_admin is not None:
            admin_raw = row.get(col_sku_admin)
            row_uses_admin_fallback = _shipment_admin_sku_falls_back_to_primary(admin_raw)

        raw_sku_cell = _resolve_shipment_row_sku_cell(row, col_sku_admin, col_sku_primary)
        ex_norm = normalize_shipment_excel_sku(raw_sku_cell)
        if not ex_norm:
            continue

        if row_uses_admin_fallback:
            # 어드민이 비었거나 미사용인 행은 DB SKU 매핑을 강제하지 않는다.
            # 대신 원본 상품코드를 표시에 보존할 수 있도록 내부 가상 키로 저장한다.
            canon = normalize_item_code(_make_no_admin_storage_sku(ex_norm))
            if not canon:
                continue
            raw_mkt = row.get(col_mkt) if col_mkt is not None else None
            raw_seg = row.get(col_segment) if col_segment is not None else None
            row_mkt = "" if raw_mkt is None or (isinstance(raw_mkt, float) and pd.isna(raw_mkt)) else str(raw_mkt).strip()
            row_seg = "" if raw_seg is None or (isinstance(raw_seg, float) and pd.isna(raw_seg)) else str(raw_seg).strip()
        else:
            canon = normalize_item_code(ex_norm)
            if not canon or canon not in sku_map:
                missing_sku_rows[str(ex_norm).strip()].add(excel_row)
                continue
            row_mkt = ""
            row_seg = ""

        raw_dt = row.get(col_date)
        d_only = _parse_shipment_order_date(raw_dt)
        if d_only is None:
            continue
        data_y = d_only.year
        mo_key = str(d_only.month)
        day_key = d_only.isoformat()

        q = _parse_int_qty(row.get(col_qty))
        if q is None or q == 0:
            continue

        for ch in target_channels:
            if not ch:
                continue
            bucket = acc[data_y][ch][canon]
            if row_uses_admin_fallback:
                # 어드민 미등록 행은 DB를 거치지 않으므로, 엑셀의 메타를 그대로 보존한다.
                if row_mkt and not str(bucket.get("mkt_priority") or "").strip():
                    bucket["mkt_priority"] = row_mkt
                if row_seg and not str(bucket.get("category") or "").strip():
                    bucket["category"] = row_seg
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

    out_by_year: dict[int, dict[str, list[dict[str, Any]]]] = {}
    for data_y in sorted(acc.keys()):
        year_acc = acc[data_y]
        out: dict[str, list[dict[str, Any]]] = {}
        for ch_name, by_sku in year_acc.items():
            rows_ch: list[dict[str, Any]] = []
            for sku_key, packs in by_sku.items():
                rows_ch.append(
                    {
                        "sku": sku_key,
                        "mkt_priority": str(packs.get("mkt_priority") or "").strip(),
                        "category": str(packs.get("category") or "").strip(),
                        "monthly": {str(k): int(v) for k, v in packs["monthly"].items()},
                        "daily": {str(k): int(v) for k, v in packs["daily"].items()},
                    }
                )
            if rows_ch:
                out[ch_name] = rows_ch
        if out:
            out_by_year[data_y] = out

    if not out_by_year:
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

    years_returned = sorted(out_by_year.keys())
    return out_by_year, years_returned


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
        no_admin_primary = _parse_no_admin_storage_sku(rec.sku)
        if no_admin_primary is not None:
            brand = ""
            pname = f"상품코드 {no_admin_primary}, 어드민 등록 X 상품"
            db_mkt = ""
            db_seg = ""
            display_sku = "어드민 등록 X"
        else:
            _b, _n, db_mkt, db_seg = sku_map.get(rec.sku, ("", "", "", ""))
            brand, pname = _b, _n
            display_sku = rec.sku
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
            "sku": display_sku,
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
