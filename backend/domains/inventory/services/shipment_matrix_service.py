"""출고 현황: 칩(시트명)별 품번(대표코드) 매트릭스 엑셀 파싱·조회."""

from __future__ import annotations

import io
import os
import re
import unicodedata
from collections import defaultdict
from datetime import date, datetime, timedelta
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

# 업로드 파일명: 예) 2026년_출고_현황.xlsx / .xls
_SHIPMENT_MATRIX_FILENAME_RE = re.compile(r"^(?P<year>\d{4})년_출고_현황(?:\.(?:xlsx|xls))?$", re.IGNORECASE)

_WEEKDAYS_KO = ("월", "화", "수", "목", "금", "토", "일")


def parse_shipment_matrix_filename(filename: str) -> int:
    """업로드 파일명에서 데이터 연도 추출. 형식: 「YYYY년_출고_현황.xlsx」."""
    base = os.path.basename((filename or "").strip())
    m = _SHIPMENT_MATRIX_FILENAME_RE.match(base)
    if not m:
        raise HTTPException(
            status_code=400,
            detail=(
                "파일 이름은 「YYYY년_출고_현황.xlsx」 형식이어야 합니다. "
                "예: 2026년_출고_현황.xlsx\n"
                f"현재 파일 이름: {base or '(이름 없음)'}"
            ),
        )
    y = int(m.group("year"))
    if y < 2000 or y > 2100:
        raise HTTPException(status_code=400, detail=f"연도는 2000~2100 범위여야 합니다. (파일명 기준: {y})")
    return y


def list_shipment_matrix_years(db: Session) -> list[int]:
    stmt = select(ShipmentMatrixRow.data_year).distinct()
    found = [int(x) for x in db.execute(stmt).scalars().all() if x is not None]
    return sorted(set(found))


def _norm_header_cell(raw: object) -> str:
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = unicodedata.normalize("NFKC", str(raw).strip())
    return re.sub(r"\s+", "", s).casefold()


def _display_dash(val: str | None) -> str:
    s = (val or "").strip()
    return s if s else "–"


def _ko_weekday_label(d: date) -> str:
    return _WEEKDAYS_KO[d.weekday()]


def format_matrix_day_header(d: date) -> str:
    return f"{d.month}/{d.day}({_ko_weekday_label(d)})"


def _excel_serial_to_date(serial: float) -> date:
    whole = int(serial)
    return date(1899, 12, 30) + timedelta(days=whole)


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


def _is_sku_column_header(norm: str) -> bool:
    if "품번(대표코드)" in norm:
        return True
    if norm == "대표코드":
        return True
    return "품번" in norm and "대표코드" in norm


def _is_mkt_column_header(norm: str) -> bool:
    if "마케팅" in norm and "우선순위" in norm:
        return True
    if "mkt" in norm and "우선순위" in norm:
        return True
    return "mkt" in norm and "우선" in norm


def _find_matrix_header_row(grid: list[list[Any]]) -> int:
    for i, row in enumerate(grid):
        for cell in row:
            if _is_sku_column_header(_norm_header_cell(cell)):
                return i
    return -1


def _month_col_match(norm: str) -> int | None:
    m = re.fullmatch(r"(\d{1,2})월합계", norm)
    if not m:
        return None
    return int(m.group(1))


def _parse_day_from_header_cell(raw_cell: object, default_year: int) -> tuple[date | None, str | None]:
    """(date_iso용 date, 원본 그대로 쓸 헤더 라벨)."""
    if isinstance(raw_cell, datetime):
        d = raw_cell.date()
        return d, None
    if isinstance(raw_cell, date) and not isinstance(raw_cell, datetime):
        return raw_cell, None
    if isinstance(raw_cell, (int, float)) and not isinstance(raw_cell, bool):
        if isinstance(raw_cell, float) and pd.isna(raw_cell):
            return None, None
        # 엑셀 날짜 직렬(대략)
        fv = float(raw_cell)
        if 20000 < fv < 600000:
            d = _excel_serial_to_date(fv)
            return d, None
        # YYYYMMDD 정수
        if 19000101 < fv < 21001231:
            s = str(int(fv))
            if len(s) == 8:
                y, mo, da = int(s[0:4]), int(s[4:6]), int(s[6:8])
                try:
                    return date(y, mo, da), None
                except ValueError:
                    pass
        return None, None

    s_raw = str(raw_cell or "").strip()
    if not s_raw:
        return None, None
    s = unicodedata.normalize("NFKC", s_raw)

    s_compact = re.sub(r"\s+", "", s)
    m_iso = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", s_compact)
    if m_iso:
        try:
            return (
                date(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))),
                None,
            )
        except ValueError:
            pass

    m = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", s_compact)
    if m:
        try:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3))), None
        except ValueError:
            pass

    m2 = re.match(r"^(\d{1,2})/(\d{1,2})", s)
    if m2:
        mo, da = int(m2.group(1)), int(m2.group(2))
        try:
            d = date(default_year, mo, da)
        except ValueError:
            return None, None
        if re.search(r"\([월화수목금토일]\)", s):
            return d, s_raw.strip()
        return d, None

    return None, None


def _merge_qty_maps(base: dict[str, int], incoming: dict[str, int], *, add: bool = False) -> None:
    for k, v in incoming.items():
        if v is None:
            continue
        if add:
            base[k] = int(base.get(k, 0)) + int(v)
        else:
            base[k] = int(v)


def parse_shipment_matrix_workbook(
    db: Session,
    raw: bytes,
    filename: str = "",
    *,
    kr_sku_brand_name: dict[str, tuple[str, str]] | None = None,
    data_year: int = _DEFAULT_DATA_YEAR,
) -> dict[str, list[dict[str, Any]]]:
    """매칭되는 시트만 읽어 channel_sheet → 행 dict 목록."""
    bio = io.BytesIO(raw)
    xl = pd.ExcelFile(bio)
    names_set = set(xl.sheet_names)
    missing_required = [n for n in SHIPMENT_MATRIX_SHEET_NAMES if n not in names_set]
    if missing_required:
        raise HTTPException(
            status_code=400,
            detail=(
                f"{filename or '파일'}: 출고 현황 파일에는 다음 시트가 모두 있어야 합니다: "
                + " · ".join(SHIPMENT_MATRIX_SHEET_NAMES)
                + f"\n누락: {', '.join(missing_required)}"
            ),
        )
    sku_map = kr_sku_brand_name if kr_sku_brand_name is not None else _load_kr_sku_brand_name(db)
    out: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for sheet_name in SHIPMENT_MATRIX_SHEET_NAMES:
        df = pd.read_excel(xl, sheet_name=sheet_name, header=None, dtype=object)
        grid = df.fillna("").values.tolist()
        h_row = _find_matrix_header_row(grid)
        if h_row < 0:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{filename or '파일'}: 시트 「{sheet_name}」에서 "
                    "「품번(대표코드)」 또는 「대표코드」 열이 있는 헤더 행을 찾지 못했습니다."
                ),
            )

        header_cells = grid[h_row]
        col_sku = col_mkt = col_cat = None
        month_cols: dict[int, int] = {}
        day_cols: dict[int, str] = {}  # col_index -> YYYY-MM-DD

        for j, cell in enumerate(header_cells):
            norm = _norm_header_cell(cell)
            if not norm:
                continue
            if _is_sku_column_header(norm):
                col_sku = j
            elif _is_mkt_column_header(norm):
                col_mkt = j
            elif norm == "구분":
                col_cat = j
            else:
                mo = _month_col_match(norm)
                if mo is not None:
                    month_cols[j] = mo
                    continue
                d_pair = _parse_day_from_header_cell(cell, data_year)
                if d_pair[0] is not None:
                    d, _label_override = d_pair
                    day_cols[j] = d.isoformat()

        if col_sku is None:
            raise HTTPException(
                status_code=400,
                detail=f"{filename or '파일'}: 시트 「{sheet_name}」에 품번(대표코드) 열이 없습니다.",
            )

        missing_skus: set[str] = set()
        by_sku: dict[str, dict[str, Any]] = {}

        for r in range(h_row + 1, len(grid)):
            row = grid[r]
            if col_sku >= len(row):
                continue
            code_cell = row[col_sku]
            if code_cell is None or (isinstance(code_cell, float) and pd.isna(code_cell)):
                continue
            ex_norm = normalize_shipment_excel_sku(code_cell)
            if not ex_norm:
                continue
            # 헤더가 데이터 영역에 반복된 행 스킵
            if _is_sku_column_header(_norm_header_cell(code_cell)):
                continue

            canon = _normalize_item_code(ex_norm)
            if not canon or canon not in sku_map:
                missing_skus.add(str(ex_norm).strip())
                continue

            mkt = ""
            if col_mkt is not None and col_mkt < len(row):
                mkt = str(row[col_mkt] if row[col_mkt] is not None else "").strip()
            cat = ""
            if col_cat is not None and col_cat < len(row):
                cat = str(row[col_cat] if row[col_cat] is not None else "").strip()

            monthly: dict[str, int] = {}
            daily: dict[str, int] = {}
            for j, mo in month_cols.items():
                if j >= len(row):
                    continue
                q = _parse_int_qty(row[j])
                if q is not None:
                    monthly[str(mo)] = q
            for j, iso in day_cols.items():
                if j >= len(row):
                    continue
                q = _parse_int_qty(row[j])
                if q is not None:
                    daily[iso] = q

            if canon not in by_sku:
                by_sku[canon] = {
                    "sku": canon,
                    "mkt_priority": mkt,
                    "category": cat,
                    "monthly": {},
                    "daily": {},
                }
            acc = by_sku[canon]
            # 같은 시트·같은 SKU가 여러 행이면 월·일 키별로 수량 합산(원래 동작)
            _merge_qty_maps(acc["monthly"], monthly, add=True)
            _merge_qty_maps(acc["daily"], daily, add=True)
            if mkt:
                acc["mkt_priority"] = mkt
            if cat:
                acc["category"] = cat

        if missing_skus:
            raise HTTPException(
                status_code=400,
                detail={
                    "code": "UNKNOWN_SKUS",
                    "message": (
                        "데이터베이스에 등록되어 있지 않은 상품코드가 있습니다. "
                        "아래 코드를 SKU 관리 탭에서 한국 SKU로 먼저 등록한 뒤 다시 업로드해 주세요."
                    ),
                    "skus": sorted(missing_skus),
                },
            )

        if not by_sku:
            continue

        for rec in by_sku.values():
            out[sheet_name].append(rec)

    if not out:
        names = " · ".join(SHIPMENT_MATRIX_SHEET_NAMES[:4]) + " …"
        raise HTTPException(
            status_code=400,
            detail=(
                f"{filename or '파일'}: 읽을 수 있는 출고 시트가 없습니다. "
                f"다음 이름의 시트만 읽습니다: {names}"
            ),
        )

    return dict(out)


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
        brand, pname = sku_map.get(rec.sku, ("", ""))
        mt = {str(k): int(v) for k, v in (rec.monthly_totals or {}).items()}
        dt = {str(k): int(v) for k, v in (rec.daily_totals or {}).items()}
        for k, v in mt.items():
            sum_monthly[k] += v
        for k, v in dt.items():
            sum_daily[k] += v

        row_channel = _rec_channel(rec)
        flat: dict[str, Any] = {
            "sku": rec.sku,
            "brand": brand,
            "description": pname,
            "mkt_priority": _display_dash(rec.mkt_priority),
            # 구분: SKU 매핑(DB item)이 아니라 업로드 엑셀 「구분」열 원문(category). DB 단종-only 정규화는 item 저장에만 적용.
            "segment": _display_dash(rec.category or None),
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
