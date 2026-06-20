import io
import math
import re
from datetime import date, datetime

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy.orm import Session

from shared.sku import normalize_item_code


REQUIRED_COLUMNS = {
    "itemno",
    "quantity",
}

COLUMN_ALIASES = {
    "warehouse": {
        "warehouse",
        "warehous",
    },
    "itemno": {
        "itemno",
        "item_no",
        "상품넘버",
        "아이템넘버",
        "품목코드",
        "item",
        "상품코드",
        "productcode",
    },
    "category": {"category", "카테고리"},
    "level": {"level"},
    "quantity": {"quantity", "가용재고"},
    "description": {"description", "상품명", "품명"},
    "option": {"option", "options", "옵션", "상품옵션", "variant", "variants"},
    "supplier": {"supplier", "공급처", "vendor"},
    "limit": {"limit", "유통기한"},
    "expdt": {"expdt", "사용기한", "exp_date"},
    "date": {"date", "일자", "날짜"},
}

LEVEL_NONE_SENTINEL = "__NONE__"

_INVENTORY_OPTION_PLACEHOLDER_CF = frozenset(
    {"nan", "none", "null", "#n/a", "n/a", "nat", "-", "undefined"}
)


def _normalize_inventory_text(value: object) -> str:
    raw = str(value or "").strip()
    return re.sub(r"\s+", " ", raw)


def _is_inventory_option_placeholder(value: object) -> bool:
    try:
        if pd.isna(value):
            return True
    except TypeError:
        pass
    if isinstance(value, float) and math.isnan(value):
        return True
    s = str(value or "").strip().casefold()
    if not s:
        return True
    return s in _INVENTORY_OPTION_PLACEHOLDER_CF


def _append_option_to_inventory_description(description: object, option: object) -> str:
    """SKU 매핑 `option` 열과 동일: 상품명이 있을 때만 뒤에 공백 + 옵션을 붙인다."""
    base = _normalize_inventory_text(description)
    if _is_inventory_option_placeholder(option):
        return base
    opt = _normalize_inventory_text(option)
    if not opt:
        return base
    if not base:
        return base
    return f"{base} {opt}"


COUNTRY_PATTERNS = {
    "KR": ["kr", "korea", "korean", "한국"],
    "TW": ["tw", "taiwan", "대만", "taipei"],
    "US": ["us", "usa", "america", "미국"],
    "VN": ["vn", "vietnam", "베트남", "hanoi", "ho chi minh"],
    "SG": ["sg", "singapore", "싱가포르", "싱가폴"],
    "AU": ["au", "australia", "호주", "sydney", "melbourne"],
    "UK": ["uk", "unitedkingdom", "britain", "england", "영국", "london"],
    "AE": ["ae", "uae", "dubai", "abudhabi", "아랍에미리트"],
    "HK": ["hk", "hongkong", "hong kong", "香港", "홍콩"],
    "JP": ["jp", "japan", "일본", "tokyo", "osaka"],
    "DE": ["de", "germany", "german", "독일", "berlin"],
    "TH": ["th", "thailand", "태국", "bangkok"],
}


def _normalize_key(value: str) -> str:
    return (
        str(value)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace("(", "")
        .replace(")", "")
    )


def _standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    normalized = {_normalize_key(col): col for col in df.columns}
    rename_map: dict[str, str] = {}
    for canonical, aliases in COLUMN_ALIASES.items():
        candidates = {_normalize_key(canonical)} | {_normalize_key(x) for x in aliases}
        for candidate in candidates:
            if candidate in normalized:
                rename_map[normalized[candidate]] = canonical
                break
    return df.rename(columns=rename_map)


def _parse_date_from_filename(filename: str) -> date | None:
    name = filename.lower()
    patterns = [
        r"(20\d{2})[-_\.]?([01]\d)[-_\.]?([0-3]\d)",
        r"([01]?\d)[-_\.]([0-3]?\d)",
    ]
    for pattern in patterns:
        match = re.search(pattern, name)
        if not match:
            continue
        groups = match.groups()
        try:
            if len(groups) == 3:
                y, m, d = map(int, groups)
                return datetime(y, m, d).date()
            m, d = map(int, groups)
            return datetime(datetime.today().year, m, d).date()
        except ValueError:
            continue
    return None


def _detect_country(warehouse_value: str, filename: str) -> str:
    base = f"{warehouse_value} {filename}".lower()
    for code, patterns in COUNTRY_PATTERNS.items():
        if any(pattern in base for pattern in patterns):
            return code
    return "KR"


def _normalize_country_code(raw: str) -> str | None:
    code = str(raw or "").strip().upper()
    if not code:
        return None
    if code == "SHIPMENT":
        return "SHIPMENT"
    if not re.match(r"^[A-Z]{2,4}$", code):
        return None
    return code



def _normalize_level_value(value: object) -> str | None:
    if pd.isna(value):
        return None

    raw = str(value).strip()
    if not raw:
        return None

    match = re.search(r"(\d+)", raw)
    if not match:
        return None

    normalized = match.group(1).lstrip("0")
    return normalized or "0"


def _nullable_string_like(value: object) -> str | None:
    if pd.isna(value):
        return None
    raw = str(value).strip()
    return raw or None


def _read_stock_bytes(filename: str, raw: bytes) -> pd.DataFrame:
    if not raw:
        raise HTTPException(status_code=400, detail=f"{filename}: 파일이 비어 있습니다.")

    try:
        if filename.lower().endswith(".csv"):
            df = pd.read_csv(io.BytesIO(raw))
        else:
            df = pd.read_excel(io.BytesIO(raw), sheet_name=0)
    except Exception as exc:
        raise HTTPException(
            status_code=400, detail=f"{filename}: 파일 파싱 실패 ({exc})"
        ) from exc

    df = _standardize_columns(df)
    missing = sorted(list(REQUIRED_COLUMNS.difference(set(df.columns))))
    if missing:
        raise HTTPException(
            status_code=400,
            detail=(
                f"{filename}: 필수 컬럼 누락 {missing}. "
                "필요 컬럼: itemno(또는 상품코드), quantity(또는 가용재고)"
            ),
        )
    if "description" not in df.columns:
        # 한국 재고 파일은 상품명이 없는 경우가 있어 창고/상품코드로 대체한다.
        if "warehouse" in df.columns and df["warehouse"].notna().any():
            df["description"] = df["warehouse"]
        else:
            df["description"] = df["itemno"]
    if "supplier" not in df.columns:
        df["supplier"] = ""
    if "warehouse" not in df.columns:
        df["warehouse"] = pd.Series([None] * len(df), index=df.index, dtype="object")
    df["sku"] = df["itemno"].apply(normalize_item_code)
    df["description"] = df["description"].fillna("").astype(str).str.strip()
    if "option" in df.columns:
        df["description"] = [
            _append_option_to_inventory_description(d, o)
            for d, o in zip(df["description"], df["option"])
        ]
        df = df.drop(columns=["option"])
    df["supplier"] = df["supplier"].fillna("").astype(str).str.strip()
    df["warehouse"] = df["warehouse"].apply(_nullable_string_like)
    return df


def _read_stock_file(upload_file: UploadFile) -> pd.DataFrame:
    upload_file.file.seek(0)
    raw = upload_file.file.read()
    upload_file.file.seek(0)
    return _read_stock_bytes(upload_file.filename, raw)


def _parse_date_series(series: pd.Series) -> pd.Series:
    raw = series.copy()

    normalized = (
        raw.astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    # 1) 20260301 같은 8자리 숫자는 YYYYMMDD 우선 해석
    digits_only = normalized.str.replace(r"[^0-9]", "", regex=True)
    ymd_mask = digits_only.str.match(r"^\d{8}$", na=False)
    parsed = pd.to_datetime(digits_only.where(ymd_mask), format="%Y%m%d", errors="coerce")

    # 2) 나머지 값은 일반 파싱
    fallback = pd.to_datetime(normalized.where(~ymd_mask), errors="coerce")
    parsed = parsed.where(~parsed.isna(), fallback)

    return parsed.dt.date


def aggregate_stock_files(
    files: list[UploadFile],
    keyword: str | None = None,
    start_date: str | None = None,
    end_date: str | None = None,
    date_range: str | None = None,
    level_filter: str | None = "all",
    file_dates: list[str] | None = None,
    file_countries: list[str] | None = None,
    db: Session | None = None,
) -> tuple[list[dict], list[str]]:
    if not files:
        raise HTTPException(status_code=400, detail="재고 파일을 최소 1개 이상 업로드해주세요.")

    daily_frames: list[pd.DataFrame] = []
    for idx, upload_file in enumerate(files):
        df = _read_stock_file(upload_file)
        filename_date = _parse_date_from_filename(upload_file.filename)

        override_country = None
        if file_countries and idx < len(file_countries):
            override_country = _normalize_country_code(file_countries[idx])
            if file_countries[idx] and override_country is None:
                raise HTTPException(
                    status_code=400,
                    detail=f"{upload_file.filename}: 국가 코드는 영문 2~4자리여야 합니다.",
                )

        if override_country is not None:
            df["country"] = override_country
        else:
            warehouse_col = df["warehouse"] if "warehouse" in df.columns else ""
            if isinstance(warehouse_col, str):
                df["country"] = _detect_country(warehouse_col, upload_file.filename)
            else:
                df["country"] = warehouse_col.fillna("").astype(str).apply(
                    lambda wh: _detect_country(wh, upload_file.filename)
                )

        override_date = None
        if file_dates and idx < len(file_dates):
            raw_override = str(file_dates[idx] or "").strip()
            if raw_override:
                parsed_override = pd.to_datetime(raw_override, errors="coerce")
                if pd.isna(parsed_override):
                    raise HTTPException(
                        status_code=400,
                        detail=f"{upload_file.filename}: 설정 탭의 날짜 형식이 올바르지 않습니다.",
                    )
                override_date = parsed_override.date()

        is_kr_file = (df["country"] == "KR").all()

        if override_date is not None:
            df["date"] = override_date
        elif is_kr_file and filename_date is not None:
            # 한국 현재고는 파일명 기준일이 스냅샷 기준일이다.
            df["date"] = filename_date
        elif "date" in df.columns:
            df["date"] = _parse_date_series(df["date"])
            if df["date"].isna().all():
                raise HTTPException(
                    status_code=400,
                    detail=f"{upload_file.filename}: date/일자 컬럼 파싱 실패",
                )
        else:
            inferred_date = filename_date
            # 한국 재고는 현재고 스냅샷이므로 일자 컬럼/파일명 날짜가 없어도 허용
            if inferred_date is None and is_kr_file:
                inferred_date = datetime.today().date()
            if inferred_date is None:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{upload_file.filename}: 날짜를 추론하지 못했습니다. "
                        "파일명에 2026-03-01 또는 20260301 형식을 포함하거나 date/일자 컬럼을 추가해주세요."
                    ),
                )
            df["date"] = inferred_date

        if "level" not in df.columns:
            df["level"] = pd.Series([None] * len(df), index=df.index, dtype="object")
        df["level"] = df["level"].apply(_normalize_level_value)
        df["level"] = df["level"].where(df["level"].notna(), LEVEL_NONE_SENTINEL)

        df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0).round().astype(int)
        grouped = (
            df.groupby(["date", "sku", "description", "supplier", "level", "warehouse", "country"], as_index=False)
            .agg(
                quantity=("quantity", "sum"),
            )
            .sort_values(["date", "country", "sku", "description", "supplier", "level", "warehouse"])
        )
        daily_frames.append(grouped)

    merged = pd.concat(daily_frames, ignore_index=True)
    merged["date"] = pd.to_datetime(merged["date"])
    merged = (
        merged.groupby(["date", "sku", "description", "supplier", "level", "warehouse", "country"], as_index=False)
        .agg(
            quantity=("quantity", "sum"),
        )
        .sort_values(["country", "sku", "description", "supplier", "level", "warehouse", "date"])
    )

    if db is not None:
        from domains.item.services.item_mapping_resolve import (
            apply_db_locale_columns_to_stock_merged,
        )

        apply_db_locale_columns_to_stock_merged(db, merged)
        merged = (
            merged.groupby(["date", "sku", "description", "supplier", "level", "warehouse", "country"], as_index=False)
            .agg(quantity=("quantity", "sum"))
            .sort_values(["country", "sku", "description", "supplier", "level", "warehouse", "date"])
        )

    max_date = merged["date"].max()
    if date_range == "7d":
        merged = merged[merged["date"] >= (max_date - pd.Timedelta(days=6))]
    elif date_range == "30d":
        merged = merged[merged["date"] >= (max_date - pd.Timedelta(days=29))]

    if start_date:
        start_dt = pd.to_datetime(start_date, errors="coerce")
        if pd.isna(start_dt):
            raise HTTPException(status_code=400, detail="start_date 형식이 올바르지 않습니다.")
        merged = merged[merged["date"] >= start_dt]
    if end_date:
        end_dt = pd.to_datetime(end_date, errors="coerce")
        if pd.isna(end_dt):
            raise HTTPException(status_code=400, detail="end_date 형식이 올바르지 않습니다.")
        merged = merged[merged["date"] <= end_dt]

    if keyword:
        needle = str(keyword).strip().lower()
        merged = merged[
            merged["sku"].astype(str).str.lower().str.contains(needle, na=False)
            | merged["description"].astype(str).str.lower().str.contains(needle, na=False)
        ]

    if level_filter and level_filter.lower() != "all":
        requested_level = str(level_filter).strip().lower()
        if requested_level in {"none", "null", "empty", "non"}:
            merged = merged[merged["level"] == LEVEL_NONE_SENTINEL]
        else:
            level_key = requested_level.lstrip("0") or "0"
            merged = merged[merged["level"] == level_key]

    if merged.empty:
        return [], []

    merged["date_key"] = merged["date"].dt.strftime("%Y-%m-%d")
    pivot = merged.pivot_table(
        index=["sku", "description", "supplier", "level", "warehouse", "country"],
        columns="date_key",
        values="quantity",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    pivot["level"] = pivot["level"].replace({LEVEL_NONE_SENTINEL: None})

    date_cols = sorted(
        [
            col
            for col in pivot.columns
            if col not in {"sku", "description", "supplier", "level", "warehouse", "country"}
        ]
    )
    rows = pivot.to_dict(orient="records")
    return rows, date_cols


def inspect_stock_files(
    files: list[UploadFile],
    file_countries: list[str] | None = None,
) -> list[dict]:
    if not files:
        raise HTTPException(status_code=400, detail="재고 파일을 최소 1개 이상 업로드해주세요.")

    results: list[dict] = []
    for idx, upload_file in enumerate(files):
        df = _read_stock_file(upload_file)
        filename_date = _parse_date_from_filename(upload_file.filename)

        override_country = None
        if file_countries and idx < len(file_countries):
            override_country = _normalize_country_code(file_countries[idx])
            if file_countries[idx] and override_country is None:
                raise HTTPException(
                    status_code=400,
                    detail=f"{upload_file.filename}: 국가 코드는 영문 2~4자리여야 합니다.",
                )

        if override_country is not None:
            country = override_country
        else:
            warehouse_col = df["warehouse"] if "warehouse" in df.columns else ""
            if isinstance(warehouse_col, str):
                country = _detect_country(warehouse_col, upload_file.filename)
            else:
                country = _detect_country(
                    str(warehouse_col.fillna("").astype(str).iloc[0]) if len(df) else "",
                    upload_file.filename,
                )

        inferred_date = ""
        is_kr_file = country == "KR"
        if is_kr_file and filename_date is not None:
            inferred_date = filename_date.isoformat()
        elif "date" in df.columns:
            parsed_date = _parse_date_series(df["date"])
            valid = parsed_date.dropna()
            if not valid.empty:
                # 단일 일자 파일이 일반적이므로 최빈값을 대표 일자로 사용
                inferred_date = (
                    pd.Series(valid.astype(str))
                    .value_counts()
                    .index[0]
                )
        elif filename_date is not None:
            inferred_date = filename_date.isoformat()

        results.append(
            {
                "filename": upload_file.filename,
                "date": inferred_date,
                "country": country,
            }
        )
    return results

