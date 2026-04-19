"""출고: 「3월 출고 ALL」 단일 시트(상품코드·수량·배송일·판매처) 파싱 및 와이드 피벗."""

from __future__ import annotations

import io
import re
import unicodedata
from collections import defaultdict
from datetime import date, datetime
from typing import Any

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from app.domains.inventory.models import ProductGroup, ProductLocale
from app.domains.inventory.services.inventory_aggregate_service import _normalize_item_code

# 출고 탭(판매처 분류 결과) — 표시 순서
SHIPMENT_CHANNEL_LABELS: tuple[str, ...] = (
    "국내 B2B",
    "국내 자사몰",
    "국내 외부몰",
    "쿠팡",
    "미국",
    "대만",
    "홍콩",
    "일본",
    "싱가폴",
    "독일",
    "영국",
    "호주",
    "동남아",
    "태국",
    "휠라선",
    "무상",
)

# 국내 외부몰: 긴 키워드부터 매칭 (부분 일치)
_EXT_MALL_SUBSTR_KEYWORDS: tuple[str, ...] = tuple(
    sorted(
        (
            "SSF SHOP",
            "GS SHOP",
            "W컨셉",
            "오늘의집",
            "홈앤쇼핑",
            "스토어팜",
            "무신사",
            "지그재그",
            "에이블리",
            "신세계",
            "이마트",
            "협력사",
            "현대",
            "삼성",
            "카카오",
            "퀸잇",
            "29cm",
            "29CM",
            "11번가",
            "CJ몰",
            "G마켓",
            "롯데",
            "토스",
            "화해",
        ),
        key=len,
        reverse=True,
    )
)

_CHANNEL_SORT = {c: i for i, c in enumerate(SHIPMENT_CHANNEL_LABELS)}

# 수동 매핑: 판매처 칸이 비어 있는 행용 JSON 키(프론트와 동일)
SHIPMENT_VENDOR_OVERRIDE_EMPTY_KEY = "__EMPTY__"

HEADER_SCAN_ROWS = 60

# --- 시트: 반드시 「3월 출고 ALL」만 (공백·ALL 대소문자·전각 영문만 정규화, 그 외 접두/접미 글자 불가) ---
def _norm_sheet_token(name: str) -> str:
    s = unicodedata.normalize("NFKC", str(name).strip()).lower()
    return re.sub(r"\s+", "", s)


SHIPMENT_REQUIRED_SHEET_LABEL = "3월 출고 ALL"
_SHIPMENT_SHEET_NAME_NORM = _norm_sheet_token(SHIPMENT_REQUIRED_SHEET_LABEL)


def _is_exact_shipment_workbook_sheet(sheet_name: str) -> bool:
    return _norm_sheet_token(sheet_name) == _SHIPMENT_SHEET_NAME_NORM


def _normalize_header_cell_text(raw: object) -> str:
    """엑셀 헤더: NBSP·전각 공백 등 제거 후 비교용 문자열."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return ""
    s = unicodedata.normalize("NFKC", str(raw).strip())
    return re.sub(r"\s+", "", s).casefold()


# --- 상품코드(엑셀) ---
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


def _load_own_mall_brand_bases(db: Session) -> frozenset[str]:
    """SKU 수기/매핑과 동일 출처: ProductGroup.brand 고유값. 판매처가 브랜드/브랜드2/브랜드(ARS)와 정확히 일치할 때만 자사몰."""
    brands = (
        db.execute(
            select(ProductGroup.brand).where(
                ProductGroup.brand.isnot(None),
                ProductGroup.brand != "",
            )
        )
        .scalars()
        .all()
    )
    return frozenset(b.strip() for b in brands if b and str(b).strip())


def _load_kr_sku_brand_name(db: Session) -> dict[str, tuple[str, str]]:
    """_normalize_item_code(sku) -> (brand, kr_display_name)."""
    locales = (
        db.execute(
            select(ProductLocale)
            .options(selectinload(ProductLocale.product_group))
            .where(ProductLocale.country_code == "KR")
        )
        .scalars()
        .all()
    )
    out: dict[str, tuple[str, str]] = {}
    for loc in locales:
        key = _normalize_item_code(loc.sku)
        if not key:
            continue
        pg = loc.product_group
        brand = (pg.brand or "").strip() if pg else ""
        name = (loc.name or "").strip()
        if not name and pg:
            name = (pg.kr_name or "").strip()
        out[key] = (brand, name)
    return out


def _shipment_source_loc(
    *,
    filename: str = "",
    sheet_name: str = "",
    excel_row_1based: int | None = None,
    column_label: str = "",
) -> str:
    """pandas read_excel(header=0) 기준 엑셀 시트의 1-based 행 번호."""
    if not filename and not sheet_name and excel_row_1based is None and not column_label:
        return ""
    parts: list[str] = []
    if filename:
        parts.append(f"파일 {filename}")
    if sheet_name:
        parts.append(f"시트 「{sheet_name}」")
    if excel_row_1based is not None:
        parts.append(f"{excel_row_1based}행(엑셀 기준)")
    if column_label:
        parts.append(f"열 「{column_label}」")
    return " — " + " · ".join(parts)


# --- 판매처 -> (채널탭, 창고이동 여부는 별도 플래그로 행에 저장) ---
def _try_classify_shipment_vendor_text(
    text: str,
    *,
    own_mall_brand_bases: frozenset[str] | None = None,
) -> tuple[str, bool] | None:
    """이미 strip된 판매처 문자열. 규칙에 맞으면 (채널, 창고이동여부), 아니면 None.

    own_mall_brand_bases: ProductGroup.brand 집합. 국내 자사몰은 판매처가 **정확히**
    `브랜드` / `브랜드2` / `브랜드(ARS)` 일 때만 (앞뒤 공백·창고이동 접미 제외 후 비교).
    """
    warehouse_move = "창고이동" in text

    def nfc(s: str) -> str:
        return unicodedata.normalize("NFC", s)

    t = nfc(text)
    t_own = re.sub(r"\s*창고이동\s*", "", t).strip()
    cf = t.casefold()

    for kw in ("무료", "직원구매", "체험단"):
        if kw in t:
            return "무상", warehouse_move

    if t.strip().upper() == "R" or "b2b" in cf:
        return "국내 B2B", warehouse_move

    if "쿠팡" in t:
        return "쿠팡", warehouse_move

    for kw in _EXT_MALL_SUBSTR_KEYWORDS:
        if kw.casefold() in cf:
            return "국내 외부몰", warehouse_move

    bases = own_mall_brand_bases or frozenset()
    for b in sorted(bases, key=len, reverse=True):
        if not b:
            continue
        if t_own == b or t_own == f"{b}2" or t_own == f"{b}(ARS)":
            return "국내 자사몰", warehouse_move

    region_pairs = (
        ("동남아", "동남아"),
        ("싱가폴", "싱가폴"),
        ("휠라선", "휠라선"),
        ("홍콩", "홍콩"),
        ("대만", "대만"),
        ("일본", "일본"),
        ("태국", "태국"),
        ("독일", "독일"),
        ("영국", "영국"),
        ("호주", "호주"),
        ("미국", "미국"),
    )
    for kw, label in sorted(region_pairs, key=lambda p: len(p[0]), reverse=True):
        if kw in t:
            return label, warehouse_move

    return None


def _normalize_vendor_channel_override_map(raw: dict[str, str] | None) -> dict[str, str] | None:
    if not raw:
        return None
    out: dict[str, str] = {}
    for k, v in raw.items():
        if not isinstance(k, str) or not isinstance(v, str):
            continue
        k2 = unicodedata.normalize("NFC", k.strip())
        v2 = v.strip()
        if k2:
            out[k2] = v2
    return out or None


def _vendor_override_lookup_key(tv: str) -> str:
    if not tv:
        return SHIPMENT_VENDOR_OVERRIDE_EMPTY_KEY
    return unicodedata.normalize("NFC", str(tv).strip())


def _resolve_shipment_vendor_channel(
    tv: str,
    *,
    own_mall_brand_bases: frozenset[str] | None,
    vendor_channel_overrides: dict[str, str] | None,
) -> tuple[str, bool] | None:
    classified = _try_classify_shipment_vendor_text(tv, own_mall_brand_bases=own_mall_brand_bases)
    if classified is not None:
        return classified
    if not vendor_channel_overrides:
        return None
    warehouse_move = "창고이동" in tv
    lk = _vendor_override_lookup_key(tv)
    ch = vendor_channel_overrides.get(lk)
    if ch is None or ch not in SHIPMENT_CHANNEL_LABELS:
        return None
    return (ch, warehouse_move)


def classify_shipment_vendor(
    raw: object,
    *,
    db: Session | None = None,
    filename: str = "",
    sheet_name: str = "",
    excel_row_1based: int | None = None,
    column_label: str = "판매처",
) -> tuple[str, bool]:
    """(SHIPMENT_CHANNEL_LABELS 중 하나, 창고이동 포함 여부). 단건 검증·API용."""
    loc = _shipment_source_loc(
        filename=filename,
        sheet_name=sheet_name,
        excel_row_1based=excel_row_1based,
        column_label=column_label,
    )
    text = str(raw or "").strip()
    if not text:
        raise HTTPException(status_code=400, detail=f"판매처가 비어 있는 행이 있습니다.{loc}")
    own = _load_own_mall_brand_bases(db) if db is not None else frozenset()
    got = _try_classify_shipment_vendor_text(text, own_mall_brand_bases=own)
    if got is None:
        raise HTTPException(
            status_code=400,
            detail=f"판매처를 분류할 수 없습니다: {text!r}. 매핑 규칙에 없는 값입니다.{loc}",
        )
    return got


def _parse_delivery_date(val: object) -> date:
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, pd.Timestamp):
        try:
            return val.date()
        except (ValueError, AttributeError):
            pass
    raw = str(val or "").strip()
    if not raw:
        raise ValueError("empty date")
    parsed = pd.to_datetime(raw, errors="coerce")
    if pd.isna(parsed):
        digits = re.sub(r"[^0-9]", "", raw)
        if len(digits) == 8:
            parsed = pd.to_datetime(digits, format="%Y%m%d", errors="coerce")
    if pd.isna(parsed):
        raise ValueError(raw)
    return parsed.date()


def _match_header_aliases(alias_norms: dict[str, set[str]], cells: list[object]) -> dict[str, int] | None:
    """한 행(또는 컬럼 이름 목록)에서 네 필드 열 인덱스. 모두 있으면 dict, 아니면 None."""
    mapping: dict[str, int] = {}
    for c, cell in enumerate(cells):
        norm = _normalize_header_cell_text(cell)
        if not norm:
            continue
        for canon, keys in alias_norms.items():
            if canon in mapping:
                continue
            if norm in keys:
                mapping[canon] = c
                break
    return mapping if len(mapping) == 4 else None


def _find_header_row_and_cols(df: pd.DataFrame) -> tuple[int, dict[str, int]] | None:
    """헤더 행 인덱스와 상품코드·상품수량·배송일·판매처 열 인덱스. 없으면 None.

    pandas 기본 read_excel(header=0)는 첫 행을 컬럼 이름으로만 쓰고 본문에는 넣지 않으므로,
    먼저 df.columns에서 매칭한다. 헤더가 본문에 있는 파일은 아래 행 스캔으로 처리.
    """
    aliases = {
        "상품코드": {"상품코드", "상품 코드", "품번"},
        "상품수량": {"상품수량", "수량", "상품 수량", "QTY", "Qty"},
        "배송일": {"배송일", "배송 일", "출고일", "일자"},
        "판매처": {"판매처", "채널", "거래처"},
    }
    alias_norms: dict[str, set[str]] = {}
    for canon, opts in aliases.items():
        keys = {_normalize_header_cell_text(canon)}
        for o in opts:
            keys.add(_normalize_header_cell_text(o))
        alias_norms[canon] = keys

    m = _match_header_aliases(alias_norms, list(df.columns))
    if m is not None:
        return -1, m

    n_r = min(HEADER_SCAN_ROWS, len(df))
    for r in range(n_r):
        row_cells = [df.iat[r, c] for c in range(df.shape[1])]
        m = _match_header_aliases(alias_norms, row_cells)
        if m is not None:
            return r, m
    return None


def _read_upload(upload_file: UploadFile) -> bytes:
    upload_file.file.seek(0)
    raw = upload_file.file.read()
    upload_file.file.seek(0)
    if not raw:
        raise HTTPException(status_code=400, detail=f"{upload_file.filename}: 파일이 비어 있습니다.")
    return raw


def _shipment_vendor_cell_str(vc: object) -> str:
    if vc is None or (isinstance(vc, float) and pd.isna(vc)):
        return ""
    return str(vc).strip()


def _shipment_workbook_raw_context(
    db: Session,
    raw: bytes,
    filename: str = "",
    *,
    kr_sku_brand_name: dict[str, tuple[str, str]] | None = None,
    own_mall_brand_bases: frozenset[str] | None = None,
) -> tuple[str, list[tuple[int, str, object, object, object]], dict[str, tuple[str, str]], frozenset[str]]:
    """시트·데이터 행·KR SKU 맵·자사몰 브랜드. 상품코드 미등록 시 HTTPException."""
    if not raw:
        raise HTTPException(status_code=400, detail=f"{filename}: 파일이 비어 있습니다.")
    try:
        book = pd.read_excel(io.BytesIO(raw), sheet_name=None, dtype=object, engine="openpyxl")
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"{filename}: 엑셀 파싱 실패 ({exc})") from exc

    target_name: str | None = None
    for sn in book:
        if _is_exact_shipment_workbook_sheet(str(sn)):
            target_name = str(sn)
            break
    if target_name is None:
        raise HTTPException(
            status_code=400,
            detail=(
                f"{filename}: 출고용 시트 이름은 반드시 「{SHIPMENT_REQUIRED_SHEET_LABEL}」만 사용할 수 있습니다. "
                f"(공백·ALL 대소문자·전각 문자만 다른 표기는 동일로 봅니다. "
                f"파일에 있는 시트: {', '.join(str(s) for s in book)})"
            ),
        )

    df = book[target_name]
    if not isinstance(df, pd.DataFrame) or df.empty:
        raise HTTPException(status_code=400, detail=f"{filename}: 시트「{target_name}」에 데이터가 없습니다.")

    found = _find_header_row_and_cols(df)
    if found is None:
        hint = (
            "첫 행이 헤더인 일반 엑셀은 열 이름(상품코드·상품수량·배송일·판매처 등)으로 인식합니다. "
            "그 위에 제목 행만 여러 줄 있으면, 헤더가 들어 있는 행까지 상단 60행 안에 있어야 합니다."
        )
        raise HTTPException(
            status_code=400,
            detail=f"{filename}: 헤더 행에서 상품코드·상품수량·배송일·판매처 네 열을 찾지 못했습니다. {hint}",
        )
    header_row, hdr = found
    c_code = hdr["상품코드"]
    c_qty = hdr["상품수량"]
    c_date = hdr["배송일"]
    c_vendor = hdr["판매처"]

    data_start = header_row + 1
    sku_map = kr_sku_brand_name if kr_sku_brand_name is not None else _load_kr_sku_brand_name(db)
    own_mall_brands = own_mall_brand_bases if own_mall_brand_bases is not None else _load_own_mall_brand_bases(db)
    raw_rows: list[tuple[int, str, object, object, object]] = []

    for r in range(data_start, len(df)):
        code_cell = df.iat[r, c_code]
        if pd.isna(code_cell):
            continue
        ex_norm = normalize_shipment_excel_sku(code_cell)
        if not ex_norm:
            continue
        qty_cell = df.iat[r, c_qty]
        date_cell = df.iat[r, c_date]
        vendor_cell = df.iat[r, c_vendor]
        if pd.isna(qty_cell) and pd.isna(date_cell) and pd.isna(vendor_cell):
            continue
        raw_rows.append((r + 1, ex_norm, qty_cell, date_cell, vendor_cell))

    missing_skus: set[str] = set()
    for _ln, ex_norm, *_ in raw_rows:
        k = _normalize_item_code(ex_norm)
        if not k or k not in sku_map:
            missing_skus.add(ex_norm)

    if missing_skus:
        lines = [
            "다음 상품코드는 DB(item_mapping · KR)에 없습니다. SKU 관리에 등록한 뒤 다시 업로드해 주세요.",
            *[f"- {c}" for c in sorted(missing_skus)],
        ]
        raise HTTPException(status_code=400, detail=lines)

    return target_name, raw_rows, sku_map, own_mall_brands


def _shipment_vendor_unmapped_payload(file_gaps: list[dict[str, Any]]) -> dict[str, Any]:
    """file_gaps: [{unmapped: dict[str, list[int]], empty_vendor_rows: list[int], ...}]. 행·파일 위치는 응답에 넣지 않음."""
    vendors: set[str] = set()
    has_empty_vendor = False
    for fg in file_gaps:
        umap = fg.get("unmapped") or {}
        vendors.update(str(v) for v in umap if v)
        if fg.get("empty_vendor_rows"):
            has_empty_vendor = True
    return {
        "code": "SHIPMENT_VENDOR_UNMAPPED",
        "message": "일부 판매처를 자동 분류하지 못했습니다. 화면에서 출고 탭을 지정해 주세요.",
        "channels": list(SHIPMENT_CHANNEL_LABELS),
        "unmapped": [{"vendor": v} for v in sorted(vendors)],
        "has_empty_vendor": has_empty_vendor,
    }


def scan_shipment_workbook_vendor_gaps(
    db: Session,
    raw: bytes,
    filename: str,
    *,
    kr_sku_brand_name: dict[str, tuple[str, str]],
    own_mall_brand_bases: frozenset[str],
) -> tuple[str, dict[str, list[int]], list[int]] | None:
    """판매처가 모두 규칙으로 분류되고 빈 칸이 없으면 None."""
    target_name, raw_rows, _sku_map, own_mall_brands = _shipment_workbook_raw_context(
        db,
        raw,
        filename,
        kr_sku_brand_name=kr_sku_brand_name,
        own_mall_brand_bases=own_mall_brand_bases,
    )
    vendor_empty_rows: list[int] = []
    vendor_unmapped: dict[str, list[int]] = defaultdict(list)
    for _ln, _ex_norm, _qty_cell, _date_cell, vendor_cell in raw_rows:
        excel_row_1based = _ln + 1
        tv = _shipment_vendor_cell_str(vendor_cell)
        if not tv:
            vendor_empty_rows.append(excel_row_1based)
        elif _try_classify_shipment_vendor_text(tv, own_mall_brand_bases=own_mall_brands) is None:
            vendor_unmapped[tv].append(excel_row_1based)
    if not vendor_empty_rows and not vendor_unmapped:
        return None
    return (target_name, {k: v for k, v in vendor_unmapped.items()}, sorted(set(vendor_empty_rows)))


def preflight_shipment_files_vendor_gaps(
    db: Session,
    file_blobs: list[tuple[str, bytes]],
    *,
    kr_sku_brand_name: dict[str, tuple[str, str]],
    own_mall_brand_bases: frozenset[str],
) -> None:
    """여러 파일을 한 요청으로 올릴 때 미매핑 판매처를 한 번에 알림(엑셀을 두 번 읽음)."""
    frags: list[dict[str, Any]] = []
    for fn, raw in file_blobs:
        gap = scan_shipment_workbook_vendor_gaps(
            db,
            raw,
            fn,
            kr_sku_brand_name=kr_sku_brand_name,
            own_mall_brand_bases=own_mall_brand_bases,
        )
        if gap:
            sheet, vu, ve = gap
            frags.append({"filename": fn, "sheet": sheet, "unmapped": vu, "empty_vendor_rows": ve})
    if frags:
        raise HTTPException(status_code=422, detail=_shipment_vendor_unmapped_payload(frags))


def parse_march_shipment_all_workbook(
    db: Session,
    raw: bytes,
    filename: str = "",
    *,
    kr_sku_brand_name: dict[str, tuple[str, str]] | None = None,
    own_mall_brand_bases: frozenset[str] | None = None,
    vendor_channel_overrides: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    """「3월 출고 ALL」 시트만 읽어 DB 검증·보강된 롱 레코드 반환.

    kr_sku_brand_name: 여러 파일 연속 파싱 시 `_load_kr_sku_brand_name(db)` 결과를 넘기면 DB 조회를 생략한다.
    own_mall_brand_bases: `_load_own_mall_brand_bases(db)` 캐시(요청당 1회).
    vendor_channel_overrides: 규칙에 없는 판매처·빈 칸을 `SHIPMENT_CHANNEL_LABELS` 값으로 수동 지정.
    """
    target_name, raw_rows, sku_map, own_mall_brands = _shipment_workbook_raw_context(
        db,
        raw,
        filename,
        kr_sku_brand_name=kr_sku_brand_name,
        own_mall_brand_bases=own_mall_brand_bases,
    )
    norm_overrides = _normalize_vendor_channel_override_map(vendor_channel_overrides)

    vendor_empty_rows: list[int] = []
    vendor_unmapped: dict[str, list[int]] = defaultdict(list)
    for _ln, _ex_norm, _qty_cell, _date_cell, vendor_cell in raw_rows:
        excel_row_1based = _ln + 1
        tv = _shipment_vendor_cell_str(vendor_cell)
        if not tv:
            vendor_empty_rows.append(excel_row_1based)
        elif _try_classify_shipment_vendor_text(tv, own_mall_brand_bases=own_mall_brands) is None:
            vendor_unmapped[tv].append(excel_row_1based)

    if vendor_empty_rows or vendor_unmapped:
        if norm_overrides is None:
            raise HTTPException(
                status_code=422,
                detail=_shipment_vendor_unmapped_payload(
                    [
                        {
                            "filename": filename,
                            "sheet": target_name,
                            "unmapped": {k: v for k, v in vendor_unmapped.items()},
                            "empty_vendor_rows": sorted(set(vendor_empty_rows)),
                        },
                    ],
                ),
            )
        bad_ch = sorted({c for c in norm_overrides.values() if c not in SHIPMENT_CHANNEL_LABELS})
        if bad_ch:
            raise HTTPException(
                status_code=400,
                detail=f"vendor_channel_overrides에 알 수 없는 탭이 있습니다: {bad_ch}. 허용: {list(SHIPMENT_CHANNEL_LABELS)}",
            )
        missing: list[str] = []
        for v in sorted(vendor_unmapped.keys()):
            if _vendor_override_lookup_key(v) not in norm_overrides:
                missing.append(v)
        if vendor_empty_rows and SHIPMENT_VENDOR_OVERRIDE_EMPTY_KEY not in norm_overrides:
            missing.append(SHIPMENT_VENDOR_OVERRIDE_EMPTY_KEY)
        if missing:
            raise HTTPException(
                status_code=400,
                detail={
                    "code": "SHIPMENT_VENDOR_OVERRIDE_INCOMPLETE",
                    "missing_keys": missing,
                    "message": "수동 매핑에 누락된 판매처가 있습니다.",
                },
            )

    records: list[dict[str, Any]] = []
    for _ln, ex_norm, qty_cell, date_cell, vendor_cell in raw_rows:
        canon = _normalize_item_code(ex_norm)
        brand, pname = sku_map[canon]
        # read_excel(header=0): df 행 인덱스 r = _ln-1 → 엑셀 1-based 데이터 행 = r+2 = _ln+1
        excel_row_1based = _ln + 1
        try:
            d = _parse_delivery_date(date_cell)
        except ValueError:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"배송일을 해석할 수 없습니다 ({date_cell!r})."
                    f"{_shipment_source_loc(filename=filename, sheet_name=target_name, excel_row_1based=excel_row_1based, column_label='배송일')}"
                ),
            ) from None
        qty = pd.to_numeric(qty_cell, errors="coerce")
        if pd.isna(qty):
            raise HTTPException(
                status_code=400,
                detail=(
                    f"상품수량이 숫자가 아닙니다 (상품코드 {ex_norm})."
                    f"{_shipment_source_loc(filename=filename, sheet_name=target_name, excel_row_1based=excel_row_1based, column_label='상품수량')}"
                ),
            )
        qty_i = int(round(float(qty)))
        tv = _shipment_vendor_cell_str(vendor_cell)
        classified = _resolve_shipment_vendor_channel(
            tv,
            own_mall_brand_bases=own_mall_brands,
            vendor_channel_overrides=norm_overrides,
        )
        if classified is None:
            raise HTTPException(
                status_code=500,
                detail="내부 오류: 판매처 검증 단계와 불일치합니다.",
            )
        channel, wh_move = classified
        level = "창고이동" if wh_move else None
        records.append(
            {
                "sku": canon,
                "description": pname,
                "brand": brand,
                "country": "KR",
                "channel": channel,
                "date": d,
                "quantity": qty_i,
                "level": level,
            }
        )

    if not records:
        raise HTTPException(status_code=400, detail=f"{filename}: 읽을 수 있는 출고 행이 없습니다.")
    return records


def build_shipment_wide_dashboard(long_recs: list[dict[str, Any]]) -> dict[str, Any]:
    """롱 레코드 -> 탭별 필터용 channel 필드 + 날짜 와이드 열. 셀: 숫자 또는 'n · 창고이동 m' 문자열."""
    if not long_recs:
        return {
            "summary": {"item_count": 0, "date_count": 0},
            "countries": ["KR"],
            "dates": [],
            "channels": list(SHIPMENT_CHANNEL_LABELS),
            "rows": [],
        }

    bucket: dict[tuple[str, str], dict[str, dict[str, int]]] = defaultdict(
        lambda: defaultdict(lambda: {"sale": 0, "wh": 0})
    )
    meta: dict[tuple[str, str], dict[str, Any]] = {}

    all_dates: set[str] = set()
    for r in long_recs:
        sku = str(r["sku"])
        ch = str(r["channel"])
        dk = r["date"].isoformat() if isinstance(r["date"], date) else str(r["date"])[:10]
        qty = int(r["quantity"] or 0)
        is_wh = r.get("level") == "창고이동"
        k = (sku, ch)
        if k not in meta:
            meta[k] = {
                "sku": sku,
                "brand": r.get("brand") or "",
                "description": r.get("description") or "",
                "country": "KR",
                "channel": ch,
                "supplier": r.get("brand") or "",
                "mapped_kr_sku": sku,
                "mapped_kr_name": r.get("description") or "",
            }
        if is_wh:
            bucket[k][dk]["wh"] += qty
        else:
            bucket[k][dk]["sale"] += qty
        all_dates.add(dk)

    date_list = sorted(all_dates)
    rows_out: list[dict[str, Any]] = []
    for k in sorted(bucket.keys(), key=lambda x: (_CHANNEL_SORT.get(x[1], 99), x[0], x[1])):
        byd = bucket[k]
        m = dict(meta[k])
        for dk in date_list:
            p = byd.get(dk, {"sale": 0, "wh": 0})
            sale, wh = p["sale"], p["wh"]
            if sale and wh:
                m[dk] = f"{sale} · 창고이동 {wh}"
            elif wh:
                m[dk] = f"창고이동 {wh}"
            else:
                m[dk] = int(sale)
        rows_out.append(m)

    return {
        "summary": {"item_count": len(rows_out), "date_count": len(date_list)},
        "countries": ["KR"],
        "dates": date_list,
        "channels": list(SHIPMENT_CHANNEL_LABELS),
        "rows": rows_out,
    }


# --- DB 뷰 재구성(저장된 집계) ---
def build_shipment_view_payload_from_long(long_recs: list[dict[str, Any]]) -> dict[str, Any]:
    """get_shipment_view: InventoryRow/집계에서 복원한 롱 데이터로 동일 응답."""
    return build_shipment_wide_dashboard(long_recs)


def collect_shipment_long_records_from_raw(
    db: Session,
    raw: bytes,
    filename: str = "",
    *,
    kr_sku_brand_name: dict[str, tuple[str, str]] | None = None,
    own_mall_brand_bases: frozenset[str] | None = None,
    vendor_channel_overrides: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    """persist 경로용: 단일 포맷 파싱."""
    return parse_march_shipment_all_workbook(
        db,
        raw,
        filename,
        kr_sku_brand_name=kr_sku_brand_name,
        own_mall_brand_bases=own_mall_brand_bases,
        vendor_channel_overrides=vendor_channel_overrides,
    )


def aggregate_shipment_files(
    db: Session,
    files: list[UploadFile],
    *,
    vendor_channel_overrides: dict[str, str] | None = None,
) -> tuple[list[dict], list[str]]:
    """호환용(테스트 등)."""
    kr_map = _load_kr_sku_brand_name(db)
    own_mall = _load_own_mall_brand_bases(db)
    all_long: list[dict[str, Any]] = []
    for uf in files:
        raw = _read_upload(uf)
        all_long.extend(
            parse_march_shipment_all_workbook(
                db,
                raw,
                uf.filename or "",
                kr_sku_brand_name=kr_map,
                own_mall_brand_bases=own_mall,
                vendor_channel_overrides=vendor_channel_overrides,
            )
        )
    payload = build_shipment_wide_dashboard(all_long)
    return payload["rows"], payload["dates"]
