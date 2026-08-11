import io
import logging
import math
import re
import uuid
from datetime import datetime, timezone

import pandas as pd
from fastapi import HTTPException, UploadFile
from sqlalchemy import delete, func, select
from sqlalchemy.orm import Session, selectinload

from domains.item.models import ProductGroup, ProductLocale
from shared.sku import normalize_item_code
from shared.config import get_runtime_settings

logger = logging.getLogger(__name__)

COUNTRY_CODES = ["KR", "US", "TW", "HK", "JP", "SG", "DE", "UK", "AU", "AE", "VN", "TH"]
COUNTRY_FIELD_SPECS = [("KR", "kr_name", "kr_sku")]
LEGACY_DYNAMIC_SKU_COLUMNS = {
    "us_msku": ("US", "MSKU"),
    "us_asin": ("US", "ASIN"),
    "us_fbm": ("US", "FBM"),
    "us_fba": ("US", "FBA"),
    "us_sku": ("US", "SKU"),
    "tw_sku": ("TW", "SKU"),
    "hk_sku": ("HK", "SKU"),
    "jp_sku": ("JP", "SKU"),
    "jp_fba": ("JP", "FBA"),
    "jp_stoo1": ("JP", "STOO1"),
    "jp_stoo2": ("JP", "STOO2"),
    "sg_sku": ("SG", "SKU"),
    "sg_fbs1": ("SG", "FBS1"),
    "sg_fbs2": ("SG", "FBS2"),
    "sg_fbs3": ("SG", "FBS3"),
    "sg_fbs4": ("SG", "FBS4"),
    "my_fbs1": ("SG", "FBS1"),
    "my_fbs2": ("SG", "FBS2"),
    "my_fbs3": ("SG", "FBS3"),
    "my_fbs4": ("SG", "FBS4"),
    "de_sku": ("DE", "SKU"),
    "de_fba1": ("DE", "FBA1"),
    "de_fba2": ("DE", "FBA2"),
    "uk_sku": ("UK", "SKU"),
    "uk_fba1": ("UK", "FBA1"),
    "uk_fba2": ("UK", "FBA2"),
    "uk_fba3": ("UK", "FBA3"),
    "au_sku": ("AU", "SKU"),
    "au_fba1": ("AU", "FBA1"),
    "ae_sku": ("AE", "SKU"),
    "ae_fba": ("AE", "FBA"),
    "vn_sku": ("VN", "SKU"),
    "vn_live1": ("VN", "리브1"),
    "vn_live2": ("VN", "리브2"),
    "th_sku": ("TH", "SKU"),
    "th_scgjwd1": ("TH", "SCGJWD1"),
    "th_scgjwd2": ("TH", "SCGJWD2"),
    "th_scgjwd3": ("TH", "SCGJWD3"),
}
# 엑셀·UI와 동일한 국가 표기 (한글 헤더 자동 생성용)
_COUNTRY_KO_LABEL = {
    "KR": "한국",
    "US": "미국",
    "TW": "대만",
    "HK": "홍콩",
    "JP": "일본",
    "SG": "싱가/말레",
    "DE": "독일",
    "UK": "영국",
    "AU": "호주",
    "AE": "UAE",
    "VN": "동남아",
    "TH": "태국",
}
COUNTRY_ORDER = {country_code: index for index, country_code in enumerate(COUNTRY_CODES)}
_COUNTRY_HEADER_ALIASES = {
    "AE": ("아랍", "아랍에미리트", "AE", "UAE", "uae", "dubai", "abudhabi"),
    "SG": (
        "싱가/말레",
        "싱가폴",
        "싱가포르",
        "말레이시아",
        "싱가폴/말레이시아",
        "SG",
        "MY",
        "singapore",
        "malaysia",
    ),
}
GENERIC_MAPPING_COLUMNS = ("mapping_country_code", "mapping_sku_type", "mapping_sku", "mapping_name")


def _canonical_country_code(raw: object) -> str:
    text = str(raw or "").strip()
    code = text.upper()
    norm = _normalize_mapping_header(text)
    if code == "MY" or norm in {_normalize_mapping_header(alias) for alias in _COUNTRY_HEADER_ALIASES.get("SG", ())}:
        return "SG"
    return code


def _unique_in_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for value in values:
        if value in seen:
            continue
        seen.add(value)
        out.append(value)
    return out


MAPPING_TEMPLATE_COLUMNS = _unique_in_order(
    [column for _, name_key, sku_key in COUNTRY_FIELD_SPECS for column in (name_key, sku_key)]
    + list(GENERIC_MAPPING_COLUMNS)
)
MAPPING_COUNTRY_NAME_COLUMNS = [f"{code.lower()}_name" for code in COUNTRY_CODES if code != "KR"]
MAPPING_INTERNAL_COLUMNS = _unique_in_order(MAPPING_TEMPLATE_COLUMNS + MAPPING_COUNTRY_NAME_COLUMNS)
OPTION_COLUMN_CANONICAL = "option"
OPTION_HEADER_ALIASES = frozenset({"option", "options", "옵션", "상품옵션", "variant", "variants"})
BRAND_COLUMN_CANONICAL = "brand"
BRAND_HEADER_ALIASES = frozenset({"brand", "브랜드"})
BARCODE_COLUMN_CANONICAL = "barcode"
BARCODE_HEADER_ALIASES = frozenset({"barcode", "바코드", "bar_code", "ean", "gtin"})
CATEGORY_COLUMN_CANONICAL = "category"
CATEGORY_HEADER_ALIASES = frozenset({"category", "카테고리", "상품카테고리", "상품 카테고리"})
REPRESENTATIVE_CODE_COLUMN_CANONICAL = "representative_code"
REPRESENTATIVE_CODE_HEADER_ALIASES = frozenset(
    {"representative_code", "representativecode", "대표코드", "rep_code", "repcode"}
)
SEGMENT_COLUMN_CANONICAL = "segment"
SEGMENT_HEADER_ALIASES = frozenset({"segment", "구분", "상품구분", "분류", "상품분류"})
MKT_PRIORITY_COLUMN_CANONICAL = "mkt_priority"
# `_normalize_mapping_header`: 공백·`_`·`-` 제거, casefold → 엑셀 헤더 변형과 맞춤
MKT_PRIORITY_HEADER_ALIASES = frozenset(
    {
        "mktpriority",
        "mkt",
        "마케팅우선순위",
        "마케팅우선",
        "마케팅등급",
    }
)
OPTIONAL_MAPPING_COLUMNS = [
    OPTION_COLUMN_CANONICAL,
    BRAND_COLUMN_CANONICAL,
    BARCODE_COLUMN_CANONICAL,
    CATEGORY_COLUMN_CANONICAL,
    REPRESENTATIVE_CODE_COLUMN_CANONICAL,
    SEGMENT_COLUMN_CANONICAL,
    MKT_PRIORITY_COLUMN_CANONICAL,
]
MAX_MAPPING_FILE_SIZE_BYTES = 2 * 1024 * 1024
# AE/UK/AU 엑셀은 전용 name 열 대신 us_name 만 있는 경우가 많음
MAPPING_NAME_FALLBACK_TO_US_NAME = frozenset({"AU", "UK", "AE"})


def _mapping_country_column_alias_sets() -> dict[str, frozenset[str]]:
    """canonical(kr_sku 등) → 허용 헤더 집합. 템플릿 한글명·영문 코드·자주 쓰는 변형 포함."""
    out: dict[str, frozenset[str]] = {}
    for code, name_key, sku_key in COUNTRY_FIELD_SPECS:
        label = _COUNTRY_KO_LABEL[code]
        ko_sku = f"{label} SKU"
        ko_name = f"{label} 상품명"
        sku_aliases = {
            sku_key,
            ko_sku,
            f"{label}SKU",
            f"{label}상품코드",
            f"{label} 상품코드",
            f"{label}품번",
            f"{label} 품번",
            f"{code} SKU",
            f"{code.lower()} sku",
        }
        name_aliases = {
            name_key,
            ko_name,
            f"{label}상품명",
            f"{label}명",
            f"{label} 상품",
            f"{code} 상품명",
            f"{code.lower()} 상품명",
        }
        if code == "SG":
            sku_aliases |= {
                "싱가폴 SKU",
                "싱가포르 SKU",
                "말레이시아 SKU",
                "싱가폴 상품코드",
                "싱가포르 상품코드",
                "말레이시아 상품코드",
                "MY SKU",
                "my sku",
            }
            name_aliases |= {
                "싱가폴 상품명",
                "싱가포르 상품명",
                "말레이시아 상품명",
                "싱가폴명",
                "싱가포르명",
                "말레이시아명",
                "MY 상품명",
                "my 상품명",
                "my_name",
            }
        if code == "KR":
            sku_aliases |= {
                "상품코드",
                "SKU",
                "품번",
                "마스터 SKU",
                "마스터SKU",
                "마스터코드",
                "마스터 코드",
                "어드민코드",
                "어드민 코드",
                "이지어드민코드",
                "이지어드민 코드",
            }
            name_aliases |= {"상품명", "제품명", "상품 이름"}
        out[sku_key] = frozenset(sku_aliases)
        out[name_key] = frozenset(name_aliases)
    for sku_key in LEGACY_DYNAMIC_SKU_COLUMNS:
        aliases = {sku_key}
        if sku_key == "tw_sku":
            aliases |= {"대만 상품코드", "대만상품코드", "TW SKU", "tw sku"}
        elif sku_key == "hk_sku":
            aliases |= {"홍콩 상품코드", "홍콩상품코드", "HK SKU", "hk sku"}
        out[sku_key] = frozenset(aliases)
    for code in COUNTRY_CODES:
        if code == "KR":
            continue
        label = _COUNTRY_KO_LABEL[code]
        name_key = f"{code.lower()}_name"
        out[name_key] = frozenset(
            {
                name_key,
                f"{label} 상품명",
                f"{label}상품명",
                f"{label}명",
                f"{label} 상품",
                f"{code} 상품명",
                f"{code.lower()} 상품명",
            }
        )
    out.update(
        {
            "mapping_country_code": frozenset({"mapping_country_code", "country_code", "국가코드", "국가 코드"}),
            "mapping_sku_type": frozenset({"mapping_sku_type", "sku_type", "상품코드 종류", "상품코드종류"}),
            "mapping_sku": frozenset({"mapping_sku", "해외 상품코드"}),
            "mapping_name": frozenset({"mapping_name", "name", "해외 상품명"}),
        }
    )
    return out


_SKU_TRAILING_VARIANT_SUFFIX = re.compile(r"^(.+)-([A-Za-z0-9]{1,12})$")


def _sku_strip_one_variant_suffix(sku: str) -> str | None:
    if not sku:
        return None
    m = _SKU_TRAILING_VARIANT_SUFFIX.match(sku)
    if not m:
        return None
    base = m.group(1)
    return base if len(base) >= 2 else None


def _normalize_mapping_name(value: object) -> str:
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    raw = str(value or "").strip()
    return re.sub(r"\s+", " ", raw)


def _normalize_mapping_name_key(value: object) -> str:
    return _normalize_mapping_name(value).casefold()


_PLACEHOLDER_MAPPING_NAME_CF = frozenset(
    {"nan", "none", "null", "#n/a", "n/a", "nat", "-", "undefined"}
)


def _is_placeholder_mapping_name(value: object) -> bool:
    """엑셀 빈칸·pandas NaN 등이 'nan' 문자열로 들어와 DB 더미 상품과 매칭되는 것을 막음."""
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
    return s in _PLACEHOLDER_MAPPING_NAME_CF


def _normalize_mapping_sku(value: object) -> str:
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    raw = str(value or "").strip()
    if not raw:
        return ""
    return normalize_item_code(raw)


def _normalize_mapping_header(value: object) -> str:
    raw = str(value or "")
    normalized = raw.replace("\ufeff", "").replace("\u200b", "").strip().casefold()
    normalized = re.sub(r"\s+", "", normalized)
    normalized = re.sub(r"[_\-()]", "", normalized)
    return normalized


def _normalize_mapping_sku_type(value: object) -> str:
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    return re.sub(r"\s+", " ", str(value or "").strip())


def _country_code_from_dynamic_sku_header(raw_label: str) -> str:
    norm = _normalize_mapping_header(raw_label)
    matches: list[str] = []
    for code in COUNTRY_CODES:
        if code == "KR":
            continue
        label = _COUNTRY_KO_LABEL[code]
        candidates = {
            _normalize_mapping_header(code),
            _normalize_mapping_header(label),
            *(_normalize_mapping_header(alias) for alias in _COUNTRY_HEADER_ALIASES.get(code, ())),
        }
        if any(candidate and candidate in norm for candidate in candidates):
            matches.append(code)
    return matches[0] if len(matches) == 1 else ""


def _country_code_from_mapping_filename(filename: str) -> str:
    stem = re.sub(r"\.xlsx$", "", str(filename or "").strip(), flags=re.IGNORECASE)
    norm = _normalize_mapping_header(stem)
    matches: list[str] = []
    for code in COUNTRY_CODES:
        if code == "KR":
            continue
        label = _COUNTRY_KO_LABEL[code]
        candidates = {
            _normalize_mapping_header(code),
            _normalize_mapping_header(label),
            *(_normalize_mapping_header(alias) for alias in _COUNTRY_HEADER_ALIASES.get(code, ())),
        }
        if any(candidate and candidate in norm for candidate in candidates):
            matches.append(code)
    return matches[0] if len(matches) == 1 else ""


def _sku_type_from_dynamic_sku_header(raw_label: str, country_code: str = "") -> str:
    raw = str(raw_label or "").replace("\ufeff", "").replace("\u200b", "").strip()
    if not raw:
        return ""
    paren = re.search(r"[\(（]\s*([^\)）]+?)\s*[\)）]", raw)
    if paren:
        return _normalize_mapping_sku_type(paren.group(1))
    cleaned = raw
    if country_code:
        label = _COUNTRY_KO_LABEL.get(country_code, "")
        cleaned = re.sub(re.escape(country_code), " ", cleaned, flags=re.IGNORECASE)
        if label:
            cleaned = cleaned.replace(label, " ")
        for alias in _COUNTRY_HEADER_ALIASES.get(country_code, ()):
            cleaned = re.sub(re.escape(alias), " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"상품\s*코드|상품코드|코드|품번|sku", " ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"[_\-()/（）]", " ", cleaned)
    cleaned = _normalize_mapping_sku_type(cleaned)
    if not cleaned or cleaned.casefold() in {"상품코드", "코드", "sku", "품번"}:
        return ""
    return cleaned


def _dynamic_sku_header_info(raw_label: str) -> dict[str, str] | None:
    raw = str(raw_label or "").strip()
    raw_norm = _normalize_mapping_header(raw)
    has_code_marker = bool(re.search(r"상품\s*코드|상품코드|코드|품번|sku", raw, flags=re.IGNORECASE))
    bare_type_allowed = bool(re.fullmatch(r"[A-Za-z0-9가-힣]{2,16}", raw)) and raw_norm not in {
        _normalize_mapping_header(v)
        for v in ("비고", "메모", "참고", "상품명", "국가", "국가코드", "브랜드", "구분")
    }
    if not has_code_marker and not bare_type_allowed:
        return None
    country_code = _country_code_from_dynamic_sku_header(raw_label)
    sku_type = _sku_type_from_dynamic_sku_header(raw_label, country_code)
    if not sku_type:
        return None
    return {"country_code": country_code, "sku_type": sku_type}


def _build_mapping_country_header_key_to_canonical() -> dict[str, str]:
    """정규화된 헤더 키 → 내부 컬럼명. import 시 별칭 충돌이 있으면 즉시 실패한다."""
    key_to_canonical: dict[str, str] = {}
    for canonical, labels in _mapping_country_column_alias_sets().items():
        for label in labels:
            nk = _normalize_mapping_header(label)
            existing = key_to_canonical.get(nk)
            if existing is not None and existing != canonical:
                raise RuntimeError(
                    f"SKU 매핑 국가 컬럼 헤더 별칭 충돌: {label!r} → {canonical}, 이미 {existing!r}에 연결됨."
                )
            if existing is None:
                key_to_canonical[nk] = canonical
    return key_to_canonical


MAPPING_COUNTRY_HEADER_KEY_TO_CANONICAL = _build_mapping_country_header_key_to_canonical()


def _lookup_group_by_country_sku_variants(
    sku_index: dict[tuple[str, str, str], ProductGroup],
    country_code: str,
    raw_sku: object,
    sku_type: object = "",
) -> ProductGroup | None:
    sku = _normalize_mapping_sku(raw_sku)
    if not sku:
        return None
    cc = str(country_code or "").strip().upper()
    st = _normalize_mapping_sku_type(sku_type)
    cur: str | None = sku
    seen: set[str] = set()
    while cur and cur not in seen:
        seen.add(cur)
        hit = sku_index.get(_sku_index_key(cc, st, cur))
        if hit is not None:
            return hit
        if not st:
            for key, group in sku_index.items():
                if key[0] == cc and key[2] == cur:
                    return group
        nxt = _sku_strip_one_variant_suffix(cur)
        if not nxt or nxt == cur:
            break
        cur = nxt
    return None


def _tw_hk_five_digit_core_from_tw_sku(tw_sku: object) -> str | None:
    """대만·홍콩 SKU 공통: 5자리 코어가 있으면 한국 마스터의 5자리 코어와 매칭.
    HK06019 / HK06019-02 → 접두 영문 뒤 5자리 우선. 그 외는 한국 SKU 와 동일 규칙으로 추출.
    연속 숫자 5자리를 뽑을 수 없으면 None → 해당 값 전체를 그 국가 SKU 로 두고 신규 item 경로."""
    s = _normalize_mapping_sku(tw_sku)
    if not s:
        return None
    m = re.match(r"^[A-Za-z]+(\d{5})(?:-.+)?$", s)
    if m:
        return m.group(1)
    return _kr_five_digit_core_for_tw_match(tw_sku)


def _kr_five_digit_core_for_tw_match(kr_sku: object) -> str | None:
    """한국 마스터 SKU 에서 대만 HK 품번과 맞출 5자리 숫자(끝에서 우선)."""
    s = _normalize_mapping_sku(kr_sku)
    if not s:
        return None
    m = re.search(r"(\d{5})\s*$", s)
    if m:
        return m.group(1)
    found = re.findall(r"\d{5}", s)
    return found[-1] if found else None


def _build_kr_tw_core_index(groups: list[ProductGroup]) -> dict[str, set[uuid.UUID]]:
    idx: dict[str, set[uuid.UUID]] = {}
    for group in groups:
        for loc in group.locales:
            if loc.country_code != "KR":
                continue
            core = _kr_five_digit_core_for_tw_match(loc.sku)
            if core:
                idx.setdefault(core, set()).add(group.id)
    return idx


def _resolve_row_localized_name(source: dict[str, object], country_code: str, name_key: str) -> str:
    direct = _normalize_mapping_name(source.get(name_key))
    if direct:
        return direct
    if country_code in MAPPING_NAME_FALLBACK_TO_US_NAME:
        return _normalize_mapping_name(source.get("us_name"))
    return ""


def _normalize_generic_mapping_locale(source: dict[str, object], row_label: str) -> dict[str, str | None] | None:
    country_code = _canonical_country_code(source.get("mapping_country_code") or source.get("country_code"))
    sku_type = _normalize_mapping_sku_type(source.get("mapping_sku_type") or source.get("sku_type"))
    sku = _normalize_mapping_sku(source.get("mapping_sku") or source.get("sku"))
    name = _normalize_mapping_name(source.get("mapping_name") or source.get("name")) or None
    if not country_code and not sku_type and not sku and not name:
        return None
    if country_code == "KR":
        raise HTTPException(status_code=400, detail=f"{row_label}: 한국 상품코드는 기존 한국 SKU 컬럼을 사용해야 합니다.")
    if country_code not in COUNTRY_CODES:
        raise HTTPException(status_code=400, detail=f"{row_label}: 지원하지 않는 국가코드입니다: {country_code}")
    if not country_code or not sku_type or not sku:
        raise HTTPException(
            status_code=400,
            detail=f"{row_label}: 해외 매핑은 국가코드, 상품코드 종류, 상품코드가 모두 필요합니다.",
        )
    return {"country_code": country_code, "sku_type": sku_type, "name": name, "sku": sku}


def _append_option_to_product_name(name: object, option: object) -> str:
    """엑셀 `option` 열 값을 상품명 뒤에 한 칸 띄워 붙인다. 상품명이 비어 있으면 옵션만으로 이름을 만들지 않는다."""
    base = _normalize_mapping_name(name)
    if _is_placeholder_mapping_name(option):
        return base
    opt = _normalize_mapping_name(option)
    if not opt:
        return base
    if not base:
        return base
    return f"{base} {opt}"


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


def _http_detail_message(detail: object) -> str:
    if isinstance(detail, str):
        return detail
    if isinstance(detail, (list, tuple)) and detail:
        return str(detail[0])
    return str(detail)


def _coerce_mapping_dataframe(df: pd.DataFrame, filename: str) -> pd.DataFrame:
    """헤더 매칭 후 표준 컬럼만 담은 DataFrame. 매핑 국가 열이 하나도 없으면 HTTPException."""
    header_key_to_canonical = MAPPING_COUNTRY_HEADER_KEY_TO_CANONICAL

    option_header_keys = frozenset(_normalize_mapping_header(a) for a in OPTION_HEADER_ALIASES)
    brand_header_keys = frozenset(_normalize_mapping_header(a) for a in BRAND_HEADER_ALIASES)
    barcode_header_keys = frozenset(_normalize_mapping_header(a) for a in BARCODE_HEADER_ALIASES)
    category_header_keys = frozenset(_normalize_mapping_header(a) for a in CATEGORY_HEADER_ALIASES)
    rep_code_header_keys = frozenset(_normalize_mapping_header(a) for a in REPRESENTATIVE_CODE_HEADER_ALIASES)
    segment_header_keys = frozenset(_normalize_mapping_header(a) for a in SEGMENT_HEADER_ALIASES)
    mkt_header_keys = frozenset(_normalize_mapping_header(a) for a in MKT_PRIORITY_HEADER_ALIASES)

    series_by_canonical: dict[str, pd.Series] = {}
    matched_raw: dict[str, str] = {}
    option_series: pd.Series | None = None
    option_matched_raw = ""
    brand_series: pd.Series | None = None
    brand_matched_raw = ""
    barcode_series: pd.Series | None = None
    barcode_matched_raw = ""
    category_series: pd.Series | None = None
    category_matched_raw = ""
    rep_code_series: pd.Series | None = None
    rep_code_matched_raw = ""
    segment_series: pd.Series | None = None
    segment_matched_raw = ""
    mkt_series: pd.Series | None = None
    mkt_matched_raw = ""
    dynamic_sku_columns: list[dict[str, object]] = []
    dynamic_matched_raw: dict[tuple[str, str], str] = {}

    for col in df.columns:
        raw_label = str(col).strip()
        norm = _normalize_mapping_header(raw_label)
        if norm in brand_header_keys:
            if brand_series is not None:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: 브랜드 열(`brand` / `브랜드`)이 두 개 이상입니다 "
                        f"(`{brand_matched_raw}` 와 `{raw_label}`)."
                    ),
                )
            col_slice = df[col]
            if isinstance(col_slice, pd.DataFrame):
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
                )
            brand_series = col_slice.copy()
            brand_matched_raw = raw_label
            continue
        if norm in barcode_header_keys:
            if barcode_series is not None:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: 바코드 열(`barcode` / `바코드` 등)이 두 개 이상입니다 "
                        f"(`{barcode_matched_raw}` 와 `{raw_label}`)."
                    ),
                )
            col_slice = df[col]
            if isinstance(col_slice, pd.DataFrame):
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
                )
            barcode_series = col_slice.copy()
            barcode_matched_raw = raw_label
            continue
        if norm in category_header_keys:
            if category_series is not None:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: 카테고리 열(`category` / `카테고리` 등)이 두 개 이상입니다 "
                        f"(`{category_matched_raw}` 와 `{raw_label}`)."
                    ),
                )
            col_slice = df[col]
            if isinstance(col_slice, pd.DataFrame):
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
                )
            category_series = col_slice.copy()
            category_matched_raw = raw_label
            continue
        if norm in rep_code_header_keys:
            if rep_code_series is not None:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: 대표코드 열이 두 개 이상입니다 "
                        f"(`{rep_code_matched_raw}` 와 `{raw_label}`)."
                    ),
                )
            col_slice = df[col]
            if isinstance(col_slice, pd.DataFrame):
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
                )
            rep_code_series = col_slice.copy()
            rep_code_matched_raw = raw_label
            continue
        if norm in segment_header_keys:
            if segment_series is not None:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: 구분 열(`segment` / `구분` 등)이 두 개 이상입니다 "
                        f"(`{segment_matched_raw}` 와 `{raw_label}`)."
                    ),
                )
            col_slice = df[col]
            if isinstance(col_slice, pd.DataFrame):
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
                )
            segment_series = col_slice.copy()
            segment_matched_raw = raw_label
            continue
        if norm in mkt_header_keys:
            if mkt_series is not None:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: 마케팅 우선순위 열이 두 개 이상입니다 "
                        f"(`{mkt_matched_raw}` 와 `{raw_label}`)."
                    ),
                )
            col_slice = df[col]
            if isinstance(col_slice, pd.DataFrame):
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
                )
            mkt_series = col_slice.copy()
            mkt_matched_raw = raw_label
            continue
        if norm in option_header_keys:
            if option_series is not None:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: `option` 열에 해당하는 헤더가 두 열 이상입니다 "
                        f"(`{option_matched_raw}` 와 `{raw_label}`)."
                    ),
                )
            col_slice = df[col]
            if isinstance(col_slice, pd.DataFrame):
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
                )
            option_series = col_slice.copy()
            option_matched_raw = raw_label
            continue
        if norm not in header_key_to_canonical:
            dynamic_info = _dynamic_sku_header_info(raw_label)
            if dynamic_info is None:
                continue
            col_slice = df[col]
            if isinstance(col_slice, pd.DataFrame):
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
                )
            dyn_key = (
                str(dynamic_info.get("country_code") or ""),
                _normalize_mapping_sku_type(dynamic_info.get("sku_type")).casefold(),
            )
            if dyn_key in dynamic_matched_raw:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: 같은 상품코드 타입 헤더가 두 열 이상입니다 "
                        f"(`{dynamic_matched_raw[dyn_key]}` 와 `{raw_label}`)."
                    ),
                )
            dynamic_matched_raw[dyn_key] = raw_label
            dynamic_sku_columns.append(
                {
                    "raw_label": raw_label,
                    "series": col_slice.copy(),
                    "country_code": dynamic_info.get("country_code") or "",
                    "sku_type": dynamic_info.get("sku_type") or "",
                }
            )
            continue
        canonical = header_key_to_canonical[norm]
        if canonical in LEGACY_DYNAMIC_SKU_COLUMNS:
            col_slice = df[col]
            if isinstance(col_slice, pd.DataFrame):
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
                )
            country_code, sku_type = LEGACY_DYNAMIC_SKU_COLUMNS[canonical]
            dyn_key = (country_code, _normalize_mapping_sku_type(sku_type).casefold())
            if dyn_key in dynamic_matched_raw:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: 같은 상품코드 타입 헤더가 두 열 이상입니다 "
                        f"(`{dynamic_matched_raw[dyn_key]}` 와 `{raw_label}`)."
                    ),
                )
            dynamic_matched_raw[dyn_key] = raw_label
            dynamic_sku_columns.append(
                {
                    "raw_label": raw_label,
                    "series": col_slice.copy(),
                    "country_code": country_code,
                    "sku_type": sku_type,
                }
            )
            continue
        if canonical in series_by_canonical:
            prev = matched_raw.get(canonical, "")
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{filename}: 같은 의미의 헤더가 두 열 이상입니다 (`{prev}` 와 `{raw_label}` → `{canonical}`)."
                ),
            )
        col_slice = df[col]
        if isinstance(col_slice, pd.DataFrame):
            raise HTTPException(
                status_code=400,
                detail=f"{filename}: 헤더 `{raw_label}` 열이 엑셀에서 중복되어 있습니다.",
            )
        series_by_canonical[canonical] = col_slice.copy()
        matched_raw[canonical] = raw_label

    if not series_by_canonical and not dynamic_sku_columns:
        raise HTTPException(
            status_code=400,
            detail=(
                f"{filename}: 인식된 매핑 헤더가 없습니다. "
                f"다음 중 일부 열 이름이 필요합니다: {', '.join(MAPPING_TEMPLATE_COLUMNS)}."
            ),
        )
    if df.empty:
        raise HTTPException(status_code=400, detail=f"{filename}: 최소 1개 이상의 매핑 행이 필요합니다.")

    aligned: dict[str, pd.Series] = {}
    for canonical in MAPPING_INTERNAL_COLUMNS:
        if canonical in series_by_canonical:
            aligned[canonical] = series_by_canonical[canonical]
        else:
            aligned[canonical] = pd.Series([pd.NA] * len(df), index=df.index, dtype=object)
    overseas_rows: list[list[dict[str, object]]] = []
    for row_idx in df.index:
        row_locales: list[dict[str, object]] = []
        row_country_code = _canonical_country_code(_normalize_mapping_sku_type(aligned["mapping_country_code"].get(row_idx)))
        row_name = _normalize_mapping_name(aligned["mapping_name"].get(row_idx))
        for spec in dynamic_sku_columns:
            raw_sku = spec["series"].get(row_idx)
            sku = _normalize_mapping_sku(raw_sku)
            if not sku:
                continue
            country_code = _canonical_country_code(
                _country_code_from_mapping_filename(filename) or row_country_code or spec.get("country_code")
            )
            if not country_code:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{filename}: `{spec['raw_label']}` 헤더에서 국가를 알 수 없습니다. "
                        "SKU 파일명에 국가 이름을 포함하거나 `mapping_country_code` 열을 추가해 주세요."
                    ),
                )
            if country_code == "KR":
                raise HTTPException(
                    status_code=400,
                    detail=f"{filename}: 한국 상품코드는 `kr_sku` 또는 `상품코드` 열을 사용해야 합니다.",
                )
            if country_code not in COUNTRY_CODES:
                raise HTTPException(status_code=400, detail=f"{filename}: 지원하지 않는 국가코드입니다: {country_code}")
            name = _normalize_mapping_name(aligned.get(f"{country_code.lower()}_name").get(row_idx)) or row_name
            row_locales.append(
                {
                    "country_code": country_code,
                    "sku_type": spec.get("sku_type") or "",
                    "sku": raw_sku,
                    "name": name or None,
                }
            )
        overseas_rows.append(row_locales)
    aligned["overseas_locales"] = pd.Series(overseas_rows, index=df.index, dtype=object)
    aligned[OPTION_COLUMN_CANONICAL] = (
        option_series if option_series is not None else pd.Series([pd.NA] * len(df), index=df.index, dtype=object)
    )
    aligned[BRAND_COLUMN_CANONICAL] = (
        brand_series if brand_series is not None else pd.Series([pd.NA] * len(df), index=df.index, dtype=object)
    )
    aligned[BARCODE_COLUMN_CANONICAL] = (
        barcode_series if barcode_series is not None else pd.Series([pd.NA] * len(df), index=df.index, dtype=object)
    )
    aligned[CATEGORY_COLUMN_CANONICAL] = (
        category_series if category_series is not None else pd.Series([pd.NA] * len(df), index=df.index, dtype=object)
    )
    aligned[REPRESENTATIVE_CODE_COLUMN_CANONICAL] = (
        rep_code_series if rep_code_series is not None else pd.Series([pd.NA] * len(df), index=df.index, dtype=object)
    )
    aligned[SEGMENT_COLUMN_CANONICAL] = (
        segment_series if segment_series is not None else pd.Series([pd.NA] * len(df), index=df.index, dtype=object)
    )
    aligned[MKT_PRIORITY_COLUMN_CANONICAL] = (
        mkt_series if mkt_series is not None else pd.Series([pd.NA] * len(df), index=df.index, dtype=object)
    )
    return pd.DataFrame(aligned)


def _read_mapping_frame(upload_file: UploadFile) -> pd.DataFrame:
    """표준 매핑 컬럼 + 선택 `brand`·`barcode`·`option`·`segment`(구분)·`mkt_priority`(마케팅 우선순위) 만 사용. 첫 시트만."""
    raw, _ = _read_mapping_bytes(upload_file)
    filename = str(upload_file.filename or "mapping.xlsx")
    last_no_mapping_detail: str | None = None
    parse_exc: Exception | None = None
    for header_row in (0, 1):
        try:
            df = pd.read_excel(io.BytesIO(raw), sheet_name=0, dtype=object, header=header_row)
        except Exception as exc:
            parse_exc = exc
            if header_row == 0:
                raise HTTPException(status_code=400, detail=f"{filename}: 엑셀 파싱 실패 ({exc})") from exc
            break
        try:
            return _coerce_mapping_dataframe(df, filename)
        except HTTPException as e:
            msg = _http_detail_message(e.detail)
            if e.status_code == 400 and "인식된 매핑 헤더가 없습니다" in msg:
                last_no_mapping_detail = msg
                continue
            raise
    if last_no_mapping_detail:
        raise HTTPException(status_code=400, detail=last_no_mapping_detail)
    if parse_exc:
        raise HTTPException(status_code=400, detail=f"{filename}: 엑셀 파싱 실패 ({parse_exc})") from parse_exc
    raise HTTPException(status_code=400, detail=f"{filename}: 엑셀을 읽을 수 없습니다.")


def _normalize_brand_cell(value: object) -> str:
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    if isinstance(value, float) and math.isnan(value):
        return ""
    s = re.sub(r"\s+", " ", str(value or "").strip())
    return s[:255]


def _normalize_barcode_cell(value: object) -> str:
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    if isinstance(value, float) and math.isnan(value):
        return ""
    s = str(value or "").strip()
    return s[:255]


def _normalize_category_cell(value: object) -> str | None:
    try:
        if pd.isna(value):
            return None
    except TypeError:
        pass
    if isinstance(value, float) and math.isnan(value):
        return None
    s = re.sub(r"\s+", " ", str(value or "").strip())
    return s[:255] if s else None


def _normalize_segment_cell(value: object) -> str:
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    if isinstance(value, float) and math.isnan(value):
        return ""
    s = re.sub(r"\s+", " ", str(value or "").strip())
    return s[:255]


_SEGMENT_LEADING_MARK_RE = re.compile(r"^\s*[（(]\s*[xXＸｘ]\s*[）)]\s*")


def normalize_item_segment(value: object) -> str | None:
    """구분: 선행 (X) / （X） 수식어만 제거 후 문자열 그대로 저장. 빈 값은 None."""
    s = _normalize_segment_cell(value)
    if not s:
        return None
    t = _SEGMENT_LEADING_MARK_RE.sub("", s).strip()
    if not t:
        return None
    return t[:255]


def normalize_mkt_priority_cell(value: object) -> str | None:
    """마케팅 우선순위(등급): 공백 정리 후 저장."""
    try:
        if pd.isna(value):
            return None
    except TypeError:
        pass
    if isinstance(value, float) and math.isnan(value):
        return None
    s = re.sub(r"\s+", " ", str(value or "").strip())
    if not s:
        return None
    return s[:255]


def format_item_segment_display(raw: str | None) -> str:
    """구분 표시용(저장값과 동일하게 노출)."""
    if raw is None:
        return ""
    return str(raw).strip()


def _resolve_representative_code_for_mapping(explicit: object, kr_sku: str) -> str:
    raw = _normalize_mapping_sku(explicit)
    if raw:
        return raw
    return kr_sku


def _serialize_mapping_release_month(group: ProductGroup) -> str:
    if group.release_month:
        return str(group.release_month).strip()
    if group.release_date:
        return group.release_date.strftime("%Y%m")
    return ""


def _serialize_mapping_code_registered_at(group: ProductGroup) -> str:
    raw = str(group.code_registered_at or "").strip()
    if not raw:
        return ""
    if "T" in raw:
        return raw.split("T", 1)[0]
    if " " in raw and len(raw) >= 10 and raw[4:5] == "-" and raw[7:8] == "-":
        return raw.split(" ", 1)[0]
    return raw


def _mapping_item_discontinued(group: ProductGroup) -> bool:
    return normalize_item_segment(group.segment) == "단종"


def _filled_product_search_detail_count(group: ProductGroup) -> int:
    locales = group.locales or []
    kr_locale = next(
        (loc for loc in locales if str(loc.country_code or "").strip().upper() == "KR"),
        None,
    )
    kr_sku = str(kr_locale.sku if kr_locale and kr_locale.sku else "").strip()
    rep = str(group.representative_code or "").strip() or kr_sku
    seg_norm = normalize_item_segment(group.segment) or ""
    values = {
        "barcode": str(group.barcode or "").strip(),
        "category": str(group.category or "").strip(),
        "representative_code": rep,
        "version": str(group.version or "").strip(),
        "kr_name": str(group.kr_name or "").strip(),
        "segment": seg_norm,
        "stock_category": str(group.stock_category or "").strip(),
        "fcst_grade": str(group.fcst_grade or "").strip(),
        "stock_grade": str(group.stock_grade or "").strip(),
        "release_month": _serialize_mapping_release_month(group),
        "code_registered_at": _serialize_mapping_code_registered_at(group),
        "kr_grade": str(group.kr_grade or "").strip(),
        "us_grade": str(group.us_grade or "").strip(),
        "tw_grade": str(group.tw_grade or "").strip(),
        "hk_grade": str(group.hk_grade or "").strip(),
        "jp_grade": str(group.jp_grade or "").strip(),
    }
    return sum(1 for value in values.values() if value)


def _mapping_item_payload(group: ProductGroup) -> dict:
    locales = sorted(group.locales, key=lambda locale: COUNTRY_ORDER.get(locale.country_code, 999))
    seg_norm = normalize_item_segment(group.segment)
    mkt_norm = normalize_mkt_priority_cell(group.mkt_priority)
    kr_locale = next((loc for loc in locales if str(loc.country_code or "").strip().upper() == "KR"), None)
    kr_sku = str(kr_locale.sku if kr_locale and kr_locale.sku else "").strip()
    rep = str(group.representative_code or "").strip() or kr_sku
    return {
        "group_id": str(group.id),
        "kr_name": group.kr_name,
        "brand": group.brand,
        "barcode": group.barcode,
        "category": group.category,
        "representative_code": rep,
        "segment": seg_norm,
        "segment_display": format_item_segment_display(seg_norm),
        "mkt_priority": mkt_norm,
        "version": group.version or "",
        "stock_category": group.stock_category or "",
        "fcst_grade": group.fcst_grade or "",
        "stock_grade": group.stock_grade or "",
        "release_month": _serialize_mapping_release_month(group),
        "code_registered_at": _serialize_mapping_code_registered_at(group),
        "kr_grade": group.kr_grade or "",
        "us_grade": group.us_grade or "",
        "tw_grade": group.tw_grade or "",
        "hk_grade": group.hk_grade or "",
        "jp_grade": group.jp_grade or "",
        "locales": [
            {"country_code": locale.country_code, "sku_type": locale.sku_type, "name": locale.name, "sku": locale.sku}
            for locale in locales
        ],
    }


def _load_product_groups(db: Session) -> list[ProductGroup]:
    return (
        db.execute(select(ProductGroup).options(selectinload(ProductGroup.locales)).order_by(ProductGroup.kr_name.asc()))
        .scalars()
        .all()
    )


def _build_mapping_group(source: dict[str, object], row_label: str) -> dict:
    locales: list[dict[str, str | None]] = []
    opt_source = source.get(OPTION_COLUMN_CANONICAL)
    for country_code, name_key, sku_key in COUNTRY_FIELD_SPECS:
        name = _append_option_to_product_name(
            _resolve_row_localized_name(source, country_code, name_key),
            opt_source,
        )
        if name and _is_placeholder_mapping_name(name):
            name = ""
        sku = _normalize_mapping_sku(source.get(sku_key))
        if not name and not sku:
            continue
        if country_code == "KR":
            if not name or not sku:
                raise HTTPException(
                    status_code=400,
                    detail=f"{row_label}: KR 상품명과 KR SKU(상품코드)는 필수입니다.",
                )
            locales.append({"country_code": country_code, "sku_type": None, "name": name, "sku": sku})
            continue
        # 해외: 상품코드(SKU)만 있으면 매핑. 상품명은 선택(공백 허용).
        if not sku:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{row_label}: {country_code} 해외 상품코드(SKU)가 필요합니다. "
                    "해외 상품명만으로는 등록할 수 없습니다."
                ),
            )
        locales.append({"country_code": country_code, "sku_type": None, "name": name or None, "sku": sku})
    generic_locale = _normalize_generic_mapping_locale(source, row_label)
    if generic_locale:
        locales.append(generic_locale)
    for raw_locale in source.get("overseas_locales") or []:
        if isinstance(raw_locale, dict):
            generic_locale = _normalize_generic_mapping_locale(
                {
                    "mapping_country_code": raw_locale.get("country_code"),
                    "mapping_sku_type": raw_locale.get("sku_type"),
                    "mapping_sku": raw_locale.get("sku"),
                    "mapping_name": raw_locale.get("name"),
                },
                row_label,
            )
            if generic_locale:
                locales.append(generic_locale)
    if not locales:
        raise HTTPException(status_code=400, detail=f"{row_label}: 최소 1개 이상의 국가 매핑이 필요합니다.")
    kr_locale = next((locale for locale in locales if locale["country_code"] == "KR"), None)
    if kr_locale is None:
        raise HTTPException(status_code=400, detail=f"{row_label}: KR 상품명과 KR SKU(상품코드)는 필수입니다.")
    brand = _normalize_brand_cell(source.get(BRAND_COLUMN_CANONICAL))
    if not brand:
        raise HTTPException(status_code=400, detail=f"{row_label}: brand(또는 브랜드) 값은 필수입니다.")
    barcode = _normalize_barcode_cell(source.get(BARCODE_COLUMN_CANONICAL))
    category = _normalize_category_cell(source.get(CATEGORY_COLUMN_CANONICAL))
    segment = normalize_item_segment(source.get(SEGMENT_COLUMN_CANONICAL))
    mkt = normalize_mkt_priority_cell(source.get(MKT_PRIORITY_COLUMN_CANONICAL))
    rep = _resolve_representative_code_for_mapping(
        source.get(REPRESENTATIVE_CODE_COLUMN_CANONICAL),
        kr_locale["sku"],
    )
    return {
        "kr_name": kr_locale["name"],
        "locales": locales,
        "brand": brand,
        "barcode": barcode,
        "category": category,
        "representative_code": rep,
        "segment": segment,
        "mkt_priority": mkt,
    }


def _find_group_by_kr_anchor(groups: list[ProductGroup], payload: dict) -> ProductGroup | None:
    """수기 upsert: 기존 그룹은 KR SKU로만 식별한다. (country_code, sku) 전역 유일.
    이름만 겹치고 SKU가 다른 신규 상품이 기존 그룹에 잘못 붙는 것을 막는다."""
    kr_locale = next((locale for locale in payload["locales"] if locale["country_code"] == "KR"), None)
    if kr_locale is None:
        return None
    target_sku = _normalize_mapping_sku(kr_locale["sku"])
    if not target_sku:
        return None
    for group in groups:
        for locale in group.locales:
            if locale.country_code != "KR":
                continue
            if _normalize_mapping_sku(locale.sku) == target_sku:
                return group
    return None


def _validate_group_conflicts(groups: list[ProductGroup], payload: dict, exclude_group_id: uuid.UUID | None = None) -> None:
    seen_keys: set[tuple[str, str, str]] = set()
    for locale in payload["locales"]:
        key = _locale_tuple_key(locale["country_code"], locale.get("sku_type"), locale["sku"])
        if key in seen_keys:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{locale['country_code']} "
                    f"{str(locale.get('sku_type') or '').strip() or 'SKU'} `{locale['sku']}` 값이 중복됩니다."
                ),
            )
        seen_keys.add(key)

    for group in groups:
        if exclude_group_id is not None and group.id == exclude_group_id:
            continue
        for locale in group.locales:
            for incoming in payload["locales"]:
                if _locale_tuple_key(locale.country_code, locale.sku_type, locale.sku) == _locale_tuple_key(
                    incoming["country_code"], incoming.get("sku_type"), incoming["sku"]
                ):
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            f"{incoming['country_code']} "
                            f"{str(incoming.get('sku_type') or '').strip() or 'SKU'} `{incoming['sku']}` "
                            "는 이미 다른 상품 그룹에 등록되어 있습니다."
                        ),
                    )


def _merge_manual_mapping_locales(target: ProductGroup, payload: dict, now: datetime) -> None:
    """폼에 적힌 국가만 반영·추가하고, 나머지 국가 로케일은 유지한다."""
    target.kr_name = payload["kr_name"]
    target.brand = _normalize_brand_cell(payload.get("brand"))
    bc = _normalize_barcode_cell(payload.get(BARCODE_COLUMN_CANONICAL))
    target.barcode = bc if bc else None
    target.category = _normalize_category_cell(payload.get(CATEGORY_COLUMN_CANONICAL))
    seg_raw = payload.get(SEGMENT_COLUMN_CANONICAL)
    target.segment = normalize_item_segment(seg_raw) if str(seg_raw or "").strip() else None
    mkt_raw = payload.get(MKT_PRIORITY_COLUMN_CANONICAL)
    target.mkt_priority = normalize_mkt_priority_cell(mkt_raw) if str(mkt_raw or "").strip() else None
    rep_raw = payload.get("representative_code")
    kr_loc = next((loc for loc in payload["locales"] if loc["country_code"] == "KR"), None)
    kr_sku_for_rep = str(kr_loc["sku"] if kr_loc else "").strip()
    target.representative_code = _resolve_representative_code_for_mapping(rep_raw, kr_sku_for_rep)
    target.manual_updated_at = now
    target.updated_at = now

    locales_map = _locales_map_by_country_sku(target)
    for inc in payload["locales"]:
        country_code = str(inc["country_code"])
        sku_type = inc.get("sku_type")
        sku = str(inc["sku"] or "")
        lk = _locale_tuple_key(country_code, sku_type, sku)
        name = str(inc.get("name") or "").strip() or None
        existing = locales_map.get(lk)
        if existing is None:
            pl = ProductLocale(
                country_code=country_code,
                sku_type=None if country_code.strip().upper() == "KR" else sku_type,
                name=name,
                sku=sku,
                updated_at=now,
            )
            target.locales.append(pl)
            locales_map[lk] = pl
            continue
        if name:
            existing.name = name
        if inc.get("sku_type") and not existing.sku_type:
            existing.sku_type = inc.get("sku_type")
        existing.updated_at = now


_SKU_MAPPING_EXAMPLE_HINT = re.compile(r"^\s*예\s*:", re.IGNORECASE)


def _is_sku_mapping_template_hint_row(source: dict[str, object]) -> bool:
    """다운로드 템플릿 2행처럼 '예: …' 안내가 들어간 행은 병합하지 않는다."""
    for key in [*MAPPING_TEMPLATE_COLUMNS, "brand", "representative_code"]:
        s = str(source.get(key) or "").strip()
        if s and _SKU_MAPPING_EXAMPLE_HINT.match(s):
            return True
    return False


def _parse_merge_record(source: dict[str, object], row_label: str) -> dict | None:
    if _is_sku_mapping_template_hint_row(source):
        return None
    sku_locales: list[dict[str, str | None]] = []
    named_locales: list[dict[str, str]] = []
    opt_source = source.get(OPTION_COLUMN_CANONICAL)

    for country_code, name_key, sku_key in COUNTRY_FIELD_SPECS:
        name = _append_option_to_product_name(
            _resolve_row_localized_name(source, country_code, name_key),
            opt_source,
        )
        sku = _normalize_mapping_sku(source.get(sku_key))
        if not name and not sku:
            continue
        if sku:
            clean_name = None
            if name and not _is_placeholder_mapping_name(name):
                clean_name = name
            sku_locales.append({"country_code": country_code, "sku_type": None, "name": clean_name, "sku": sku})
        elif name and not _is_placeholder_mapping_name(name):
            named_locales.append({"country_code": country_code, "name": name})

    generic_locale = _normalize_generic_mapping_locale(source, row_label)
    if generic_locale:
        sku_locales.append(generic_locale)
    for raw_locale in source.get("overseas_locales") or []:
        if isinstance(raw_locale, dict):
            generic_locale = _normalize_generic_mapping_locale(
                {
                    "mapping_country_code": raw_locale.get("country_code"),
                    "mapping_sku_type": raw_locale.get("sku_type"),
                    "mapping_sku": raw_locale.get("sku"),
                    "mapping_name": raw_locale.get("name"),
                },
                row_label,
            )
            if generic_locale:
                sku_locales.append(generic_locale)

    if not sku_locales and not named_locales:
        return None
    if not sku_locales:
        raise HTTPException(status_code=400, detail=f"{row_label}: 최소 한 국가의 SKU 값은 필요합니다.")

    first_name = next(
        (
            str(locale.get("name") or "").strip()
            for locale in [*sku_locales, *named_locales]
            if str(locale.get("name") or "").strip()
        ),
        "",
    )
    first_sku = next(
        (str(locale.get("sku") or "").strip() for locale in sku_locales if str(locale.get("sku") or "").strip()), ""
    )
    anchor_kr_sku = _normalize_mapping_sku(source.get("kr_sku"))
    anchor_kr_name = _normalize_mapping_name(source.get("kr_name"))
    if _is_placeholder_mapping_name(anchor_kr_name):
        anchor_kr_name = ""
    anchor_kr_name = _append_option_to_product_name(anchor_kr_name, opt_source)
    match_us_name = _normalize_mapping_name(source.get("us_name"))
    if _is_placeholder_mapping_name(match_us_name):
        match_us_name = ""
    match_us_name = _append_option_to_product_name(match_us_name, opt_source)

    brand = _normalize_brand_cell(source.get(BRAND_COLUMN_CANONICAL))
    barcode = _normalize_barcode_cell(source.get(BARCODE_COLUMN_CANONICAL))
    category = _normalize_category_cell(source.get(CATEGORY_COLUMN_CANONICAL))
    segment = _normalize_segment_cell(source.get(SEGMENT_COLUMN_CANONICAL))
    mkt_priority = _normalize_segment_cell(source.get(MKT_PRIORITY_COLUMN_CANONICAL))
    kr_sku_val = next(
        (str(locale.get("sku") or "").strip() for locale in sku_locales if locale.get("country_code") == "KR"),
        anchor_kr_sku,
    )
    representative_code = _resolve_representative_code_for_mapping(
        source.get(REPRESENTATIVE_CODE_COLUMN_CANONICAL),
        kr_sku_val or first_sku,
    )

    return {
        "row_label": row_label,
        "kr_name": anchor_kr_name,
        "sku_locales": sku_locales,
        "named_locales": named_locales,
        "fallback_name": first_name or first_sku or "미분류 상품",
        "anchor_kr_sku": anchor_kr_sku,
        "anchor_kr_name": anchor_kr_name,
        "match_us_name": match_us_name,
        "brand": brand,
        "barcode": barcode,
        "category": category,
        "representative_code": representative_code,
        "segment": segment,
        "mkt_priority": mkt_priority,
    }


def _read_merge_records(upload_files: list[UploadFile]) -> list[dict]:
    records: list[dict] = []
    for upload_file in upload_files:
        df = _read_mapping_frame(upload_file)
        filename = str(upload_file.filename or "mapping.xlsx")
        for idx, row in df.iterrows():
            row_label = f"{filename} 행 {idx + 2}"
            record = _parse_merge_record(row.to_dict(), row_label)
            if record is None:
                continue
            records.append(record)
    if not records:
        raise HTTPException(status_code=400, detail="업로드한 파일에서 반영할 SKU 매핑 데이터를 찾지 못했습니다.")
    return records


def _index_key(country_code: str, value: str) -> tuple[str, str]:
    return (str(country_code or "").strip().upper(), str(value or "").strip())


def _sku_index_key(country_code: str, sku_type: object, sku: str) -> tuple[str, str, str]:
    cc = str(country_code or "").strip().upper()
    st = "" if cc == "KR" else _normalize_mapping_sku_type(sku_type).casefold()
    return (cc, st, str(sku or "").strip())


def _locale_tuple_key(country_code: str, sku_type: object, sku: str) -> tuple[str, str, str]:
    return _sku_index_key(country_code, sku_type, sku)


def _locales_map_by_country_sku(group: ProductGroup) -> dict[tuple[str, str, str], ProductLocale]:
    return {_locale_tuple_key(loc.country_code, loc.sku_type, loc.sku): loc for loc in group.locales}


def _remove_group_from_indexes(
    group: ProductGroup,
    sku_index: dict[tuple[str, str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
) -> None:
    for locale in group.locales:
        sku_key = _sku_index_key(locale.country_code, locale.sku_type, locale.sku)
        if sku_index.get(sku_key) is group:
            sku_index.pop(sku_key, None)
        if locale.name:
            name_key = _index_key(locale.country_code, _normalize_mapping_name_key(locale.name))
            group_ids = name_index.get(name_key)
            if group_ids:
                group_ids.discard(group.id)
                if not group_ids:
                    name_index.pop(name_key, None)
    group_index[group.id] = group


def _index_group(
    group: ProductGroup,
    sku_index: dict[tuple[str, str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
) -> None:
    group_index[group.id] = group
    for locale in group.locales:
        sku_index[_sku_index_key(locale.country_code, locale.sku_type, locale.sku)] = group
        if locale.name:
            name_key = _index_key(locale.country_code, _normalize_mapping_name_key(locale.name))
            name_index.setdefault(name_key, set()).add(group.id)


def _pick_primary_group(groups: list[ProductGroup], preferred_kr_name: str = "") -> ProductGroup:
    preferred_key = _normalize_mapping_name_key(preferred_kr_name)

    def _sort_key(group: ProductGroup) -> tuple[int, int, str]:
        has_preferred = int(_normalize_mapping_name_key(group.kr_name) == preferred_key and preferred_key != "")
        has_kr_locale = int(any(locale.country_code == "KR" for locale in group.locales))
        return (-has_preferred, -has_kr_locale, str(group.id))

    return sorted(groups, key=_sort_key)[0]


def _groups_matching_kr_display_name(
    anchor_kr_name: str,
    group_index: dict[uuid.UUID, ProductGroup],
) -> list[ProductGroup]:
    if not anchor_kr_name:
        return []
    key = _normalize_mapping_name_key(anchor_kr_name)
    seen: set[uuid.UUID] = set()
    out: list[ProductGroup] = []
    for group in group_index.values():
        if group.id in seen:
            continue
        if _normalize_mapping_name_key(group.kr_name) == key:
            seen.add(group.id)
            out.append(group)
            continue
        for loc in group.locales:
            if loc.country_code == "KR" and _normalize_mapping_name_key(loc.name or "") == key:
                seen.add(group.id)
                out.append(group)
                break
    return out


def _group_conflict_summary(group_index: dict[uuid.UUID, ProductGroup], gid: uuid.UUID) -> str:
    """충돌 안내용: 어떤 기존 상품인지 한눈에 대조할 수 있게 요약."""
    group = group_index.get(gid)
    if group is None:
        return f"• item_id={gid} (메모리 인덱스에 없음 — 비정상)"
    kr = (group.kr_name or "").strip() or "(한국 대표명 없음)"
    lines: list[str] = [f"• item_id={group.id}", f"  한국 대표명(kr_name): {kr!r}"]
    priority = tuple(COUNTRY_CODES)
    for cc in priority:
        for loc in sorted(
            (x for x in group.locales if x.country_code == cc),
            key=lambda x: x.sku,
        ):
            nm = (loc.name or "").strip()
            if nm:
                lines.append(f"  {cc}: SKU={loc.sku!r}, 이름={nm!r}")
            else:
                lines.append(f"  {cc}: SKU={loc.sku!r}")
    return "\n".join(lines)


def _merge_row_excerpt(record: dict) -> str:
    """충돌 진단용: 업로드 행에 실제로 들어 있는 국가·SKU·이름·앵커 열."""
    lines: list[str] = []
    for loc in record.get("sku_locales") or []:
        cc = str(loc.get("country_code") or "")
        sku = str(loc.get("sku") or "")
        nm = str(loc.get("name") or "").strip()
        if nm:
            lines.append(f"  {cc}: SKU={sku!r}, 이름={nm!r}")
        else:
            lines.append(f"  {cc}: SKU={sku!r}")
    for loc in record.get("named_locales") or []:
        cc = str(loc.get("country_code") or "")
        nm = str(loc.get("name") or "").strip()
        if nm:
            lines.append(f"  {cc}: (SKU 없음) 이름만={nm!r}")
    ak = str(record.get("anchor_kr_sku") or "").strip()
    if ak:
        lines.append(f"  엑셀 kr_sku={ak!r}")
    an = str(record.get("anchor_kr_name") or "").strip()
    if an:
        lines.append(f"  엑셀 kr_name={an!r}")
    um = str(record.get("match_us_name") or "").strip()
    if um:
        lines.append(f"  엑셀 us_name={um!r}")
    return "\n".join(lines) if lines else "  (국가·SKU 행 데이터 없음)"


def _one_line_group_hint(group_index: dict[uuid.UUID, ProductGroup], gid: uuid.UUID) -> str:
    g = group_index.get(gid)
    if not g:
        return f"id={gid}"
    kr = (g.kr_name or "").strip() or "?"
    return f"id={g.id}, kr_name={kr!r}"


def _add_matched_group_id(
    matched_ids: set[uuid.UUID],
    group_id: uuid.UUID,
    row_label: str,
    reason: str,
    group_index: dict[uuid.UUID, ProductGroup],
    record: dict,
) -> None:
    if matched_ids and group_id not in matched_ids:
        existing_blocks = [
            _group_conflict_summary(group_index, gid) for gid in sorted(matched_ids, key=str)
        ]
        new_block = _group_conflict_summary(group_index, group_id)
        clue = f"{reason}" if reason else "추가 단서"
        existing_hints = " | ".join(_one_line_group_hint(group_index, g) for g in sorted(matched_ids, key=str))
        new_hint = _one_line_group_hint(group_index, group_id)
        row_bit = _merge_row_excerpt(record)
        raise HTTPException(
            status_code=400,
            detail=(
                f"{row_label}: 행에 적힌 국가·SKU·이름이 서로 다른 기존 상품을 가리킵니다.\n"
                f"요약: 이미 확정된 상품({existing_hints}) ≠ 이번 단서가 가리키는 상품({new_hint}).\n\n"
                f"【이번 업로드 행에 적힌 값】\n{row_bit}\n\n"
                f"【이미 이 행으로 확정된 상품(상세)】\n"
                f"{chr(10).join(existing_blocks)}\n\n"
                f"【지금 맞춰 보려던 단서】 {clue}\n"
                f"【그 단서가 가리키는 상품(상세)】\n"
                f"{new_block}\n\n"
                "위 두(또는 여러) 상품이 서로 다르면 한 행에 서로 다른 item 을 넣은 것입니다. "
                "엑셀에서 국가별 SKU·이름·대만 코어/us_name/kr_sku 등이 모두 같은 상품을 가리키도록 수정해 주세요."
            ),
        )
    matched_ids.add(group_id)


def _find_matching_groups(
    record: dict,
    sku_index: dict[tuple[str, str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
    kr_tw_core_index: dict[str, set[uuid.UUID]],
) -> list[ProductGroup]:
    matched_ids: set[uuid.UUID] = set()
    row_label = str(record["row_label"])
    # 5)·6)에서 사용: 이 국가는 이미 SKU/코어 등으로 item 이 잠겼으면 상품명으로 다시 묶지 않음
    sku_resolved_cc: set[str] = set()

    # 0) 홍콩 마스터 행: hk_sku + 엑셀 kr_sku 가 있으면 kr_sku 로 기존 item 을 찾고, 없으면 신규 item 만든 뒤 HK 매핑.
    #    이 경우 TW/이름 단서로 다른 item 을 끌어오지 않는다.
    hk_kr_locked = False
    has_hk_sku = any(
        str(loc.get("country_code") or "").strip().upper() == "HK" and str(loc.get("sku") or "").strip()
        for loc in record["sku_locales"]
    )
    anchor_kr_for_hk = str(record.get("anchor_kr_sku") or "").strip()
    if has_hk_sku and anchor_kr_for_hk:
        hk_kr_locked = True
        kr_hit_hk = _lookup_group_by_country_sku_variants(sku_index, "KR", anchor_kr_for_hk)
        if kr_hit_hk is not None:
            matched_ids.add(kr_hit_hk.id)
            sku_resolved_cc.add("KR")
            sku_resolved_cc.add("HK")

    # 1) 대만·홍콩: HKS06130 등 → 5자리 코어로 한국 마스터 item 매칭 (우선).
    #    코어 추출 불가(예: HK0LDHPC1) → 해당 문자열을 그대로 TW/HK SKU 로 두고 신규 item (us_name 으로 기존 item 강제 연결 안 함).
    tw_hk_cores: list[str] = []
    has_tw_hk_sku_locale = False
    tw_hk_core_resolved = False
    tw_hk_no_core_new_item = False
    if not hk_kr_locked:
        for loc in record["sku_locales"]:
            cc = str(loc.get("country_code") or "").strip().upper()
            if cc not in ("TW", "HK"):
                continue
            if not str(loc.get("sku") or "").strip():
                continue
            has_tw_hk_sku_locale = True
            core = _tw_hk_five_digit_core_from_tw_sku(loc.get("sku"))
            if core:
                tw_hk_cores.append(core)
    if not hk_kr_locked and tw_hk_cores:
        locked_tw_hk: set[uuid.UUID] = set()
        for core in tw_hk_cores:
            hits = kr_tw_core_index.get(core, set())
            if len(hits) > 1:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{row_label}: 대만/홍콩 품번에서 뽑은 5자리 `{core}` 에 해당하는 한국 SKU 를 가진 상품이 DB 에 여러 개입니다. "
                        "한국 마스터에서 해당 5자리를 유일하게 맞춘 뒤 다시 업로드하세요."
                    ),
                )
            if len(hits) == 1:
                locked_tw_hk.add(next(iter(hits)))
        if len(locked_tw_hk) > 1:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{row_label}: 같은 행의 대만/홍콩 SKU 들이 서로 다른 한국 상품(5자리 코어)을 가리킵니다. "
                    "한 행에는 동일 상품의 코드만 넣어 주세요."
                ),
            )
        if len(locked_tw_hk) == 1:
            matched_ids.add(next(iter(locked_tw_hk)))
            sku_resolved_cc.add("TW")
            sku_resolved_cc.add("HK")
            sku_resolved_cc.add("KR")
            tw_hk_core_resolved = True
    elif not hk_kr_locked and has_tw_hk_sku_locale:
        tw_hk_no_core_new_item = True
        for loc in record["sku_locales"]:
            cc = str(loc.get("country_code") or "").strip().upper()
            if cc in ("TW", "HK") and str(loc.get("sku") or "").strip():
                sku_resolved_cc.add(cc)

    # 대만/홍콩 SKU 가 5자리 코어로 이미 한국 item 에 고정된 경우, 같은 행의 kr_name / us_name / AE·AU·UK 이름 등으로
    # 다른 item 을 끌어오면 서로 다른 상품 충돌이 나므로 아래 2)~4)·6) 이름 기반 단서는 생략한다.
    # 코어 없이 신규 item 으로 가는 TW/HK 행도 이름 단서로 기존 item 에 붙이지 않는다.
    # 홍콩(kr_sku 앵커) 행도 동일하게 이름 기반 단서를 생략한다.
    if not tw_hk_core_resolved and not hk_kr_locked and not tw_hk_no_core_new_item:
        # 2) 엑셀 kr_sku → DB 한국 마스터 앵커 (미국 파일 등)
        anchor_kr = str(record.get("anchor_kr_sku") or "").strip()
        if anchor_kr:
            kr_hit = _lookup_group_by_country_sku_variants(sku_index, "KR", anchor_kr)
            if kr_hit is not None:
                sku_resolved_cc.add("KR")
                _add_matched_group_id(
                    matched_ids,
                    kr_hit.id,
                    row_label,
                    f"엑셀 `kr_sku`={anchor_kr!r} 가 가리키는 상품과",
                    group_index,
                    record,
                )

        # 3) 베트남 등: kr_name 으로 한국 대표 상품명 매칭
        anchor_name = _normalize_mapping_name(record.get("anchor_kr_name"))
        has_vn = any(str(l.get("country_code") or "").upper() == "VN" for l in record["sku_locales"])
        if anchor_name and not _is_placeholder_mapping_name(anchor_name) and has_vn and not anchor_kr:
            name_hits = _groups_matching_kr_display_name(anchor_name, group_index)
            if len(name_hits) > 1:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{row_label}: `kr_name` `{anchor_name}` 과 일치하는 한국 상품이 여러 개입니다. "
                        "`kr_sku` 로 구분하거나 마스터 상품명을 유일하게 맞춰 주세요."
                    ),
                )
            if len(name_hits) == 1:
                sku_resolved_cc.add("KR")
                _add_matched_group_id(
                    matched_ids,
                    name_hits[0].id,
                    row_label,
                    f"`kr_name`={anchor_name!r} 이 가리키는 상품과",
                    group_index,
                    record,
                )

        # 4) AE/AU/UK: us_name 으로 DB 의 US 로케일 상품명과 매칭
        us_link = _normalize_mapping_name(record.get("match_us_name"))
        overseas_us_name_cc = frozenset({"AE", "AU", "UK"})
        needs_us_name = any(
            str(l.get("country_code") or "").strip().upper() in overseas_us_name_cc for l in record["sku_locales"]
        )
        if us_link and not _is_placeholder_mapping_name(us_link) and needs_us_name:
            candidate_ids = name_index.get(_index_key("US", _normalize_mapping_name_key(us_link)), set())
            if len(candidate_ids) > 1:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{row_label}: `us_name` `{us_link}` 과 일치하는 US 상품명을 가진 기존 상품이 여러 개입니다. "
                        "US 명을 유일하게 맞추거나 `kr_sku` 로 먼저 연결해 주세요."
                    ),
                )
            if len(candidate_ids) == 1:
                sku_resolved_cc.add("US")
                _add_matched_group_id(
                    matched_ids,
                    next(iter(candidate_ids)),
                    row_label,
                    f"`us_name`={us_link!r} 이 가리키는 상품과",
                    group_index,
                    record,
                )

    # 5) 국가별 SKU 직접 일치
    #    이후 6)에서 같은 국가는 이름으로 다시 묶지 않음 — DB에 정품/증정 등 서로 다른 item 이
    #    동일 표시명(예: 대만 번체)을 쓰는 경우 SKU·대만 코어가 우선해야 함.
    #    홍콩 행이 kr_sku 미매칭으로 신규 item 을 만들 경로일 때는, 여기서 HK SKU 가 다른 item 에 붙어
    #    잘못 합쳐지는 것을 막기 위해 5) 단계를 건너뛴다.
    skip_direct_sku_lookup = hk_kr_locked and not matched_ids
    if not skip_direct_sku_lookup:
        for locale in record["sku_locales"]:
            cc = str(locale["country_code"] or "").strip().upper()
            sku_type = locale.get("sku_type")
            sku = str(locale["sku"] or "")
            matched = _lookup_group_by_country_sku_variants(sku_index, cc, sku, sku_type)
            if matched is not None:
                sku_resolved_cc.add(cc)
                _add_matched_group_id(
                    matched_ids,
                    matched.id,
                    row_label,
                    f"{cc} SKU={sku!r} 로 DB 에서 직접 찾은 상품과",
                    group_index,
                    record,
                )

    # 6) 국가별 상품명 (대만/홍콩 5자리 코어·신규-only 행은 위에서 설명한 이유로 생략)
    if not tw_hk_core_resolved and not hk_kr_locked and not tw_hk_no_core_new_item:
        for locale in [*record["sku_locales"], *record["named_locales"]]:
            raw_nm = locale.get("name")
            if _is_placeholder_mapping_name(raw_nm):
                continue
            name = _normalize_mapping_name(raw_nm)
            if not name:
                continue
            cc_nm = str(locale.get("country_code") or "").strip().upper()
            if cc_nm in sku_resolved_cc:
                continue
            candidate_ids = name_index.get(_index_key(cc_nm, _normalize_mapping_name_key(name)), set())
            if len(candidate_ids) <= 1:
                for gid in candidate_ids:
                    _add_matched_group_id(
                        matched_ids,
                        gid,
                        row_label,
                        f"{cc_nm} 상품명={name!r} 이 가리키는 상품과",
                        group_index,
                        record,
                    )
                continue
            if matched_ids and matched_ids.intersection(candidate_ids):
                continue
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{row_label}: {cc_nm} 상품명 `{name}` 이 여러 상품과 연결되어 자동 병합할 수 없습니다."
                ),
            )

    return [group_index[gid] for gid in matched_ids if gid in group_index]


def _merge_groups(
    db: Session,
    groups: list[ProductGroup],
    preferred_kr_name: str,
    sku_index: dict[tuple[str, str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
    row_label: str,
) -> ProductGroup:
    if len(groups) <= 1:
        return groups[0]

    primary = _pick_primary_group(groups, preferred_kr_name)
    secondary_groups = [g for g in groups if g.id != primary.id]

    for group in [primary, *secondary_groups]:
        _remove_group_from_indexes(group, sku_index, name_index, group_index)

    primary_keys = _locales_map_by_country_sku(primary)

    for secondary in secondary_groups:
        for loc in list(secondary.locales):
            lk = _locale_tuple_key(loc.country_code, loc.sku_type, loc.sku)
            existing = primary_keys.get(lk)
            if existing is not None:
                if loc.name and (not existing.name or len(str(loc.name)) > len(str(existing.name or ""))):
                    existing.name = loc.name
                db.delete(loc)
                continue
            loc.product_group = primary
            primary_keys[lk] = loc
        db.delete(secondary)
        group_index.pop(secondary.id, None)

    _index_group(primary, sku_index, name_index, group_index)
    return primary


def _apply_merge_record(
    db: Session,
    target: ProductGroup,
    record: dict,
    uploaded_at: datetime,
    sku_index: dict[tuple[str, str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
) -> ProductGroup:
    _remove_group_from_indexes(target, sku_index, name_index, group_index)

    locales_map = _locales_map_by_country_sku(target)
    for locale in record["sku_locales"]:
        country_code = str(locale["country_code"])
        sku_type = locale.get("sku_type")
        sku = str(locale["sku"] or "")
        name = _normalize_mapping_name(locale.get("name"))
        lk = _locale_tuple_key(country_code, sku_type, sku)
        conflict_group = sku_index.get(_sku_index_key(country_code, sku_type, sku))
        if conflict_group is not None and conflict_group.id != target.id:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{record['row_label']}: {country_code} "
                    f"{str(sku_type or '').strip() or 'SKU'} `{sku}` 는 이미 다른 상품에 연결되어 있습니다. "
                    "충돌이 있어 전체 업로드를 중단합니다."
                ),
            )
        existing = locales_map.get(lk)
        if existing is None:
            new_locale = ProductLocale(
                country_code=country_code,
                sku_type=None if country_code.strip().upper() == "KR" else sku_type,
                name=name or None,
                sku=sku,
                updated_at=uploaded_at,
            )
            target.locales.append(new_locale)
            locales_map[lk] = new_locale
            continue
        if country_code.strip().upper() != "KR" and locale.get("sku_type") and not existing.sku_type:
            existing.sku_type = locale.get("sku_type")
        if name:
            existing.name = name
        existing.updated_at = uploaded_at
        # 이미 그룹에 (국가, SKU)가 있으면 업로드 행의 이름·시간은 반영하지 않음(신규 조합만 반영).

    for locale in record["named_locales"]:
        cc = str(locale["country_code"])
        nm = _normalize_mapping_name(locale["name"])
        for pl in target.locales:
            if pl.country_code == cc:
                pl.name = nm
                pl.updated_at = uploaded_at

    kr_rep = str(record.get("kr_name") or "").strip()
    if kr_rep:
        target.kr_name = kr_rep
    elif not str(target.kr_name or "").strip():
        target.kr_name = record["fallback_name"]
    brand_val = _normalize_brand_cell(record.get("brand"))
    if brand_val:
        target.brand = brand_val
    barcode_val = _normalize_barcode_cell(record.get(BARCODE_COLUMN_CANONICAL))
    if barcode_val:
        target.barcode = barcode_val
    category_val = _normalize_category_cell(record.get(CATEGORY_COLUMN_CANONICAL))
    if category_val:
        target.category = category_val
    seg_in = record.get("segment")
    if isinstance(seg_in, str) and seg_in.strip():
        target.segment = normalize_item_segment(seg_in)
    mkt_in = record.get("mkt_priority")
    if mkt_in is not None and str(mkt_in).strip():
        target.mkt_priority = normalize_mkt_priority_cell(mkt_in)
    rep_in = record.get("representative_code")
    if rep_in:
        target.representative_code = str(rep_in).strip()
    elif not str(target.representative_code or "").strip():
        kr_loc = next(
            (pl for pl in target.locales if str(pl.country_code or "").strip().upper() == "KR"),
            None,
        )
        if kr_loc and kr_loc.sku:
            target.representative_code = str(kr_loc.sku).strip()
    target.upload_updated_at = uploaded_at
    target.updated_at = uploaded_at

    _index_group(target, sku_index, name_index, group_index)
    return target


def get_item_sku_mapping_summary(db: Session) -> dict:
    total_count = db.execute(select(func.count(ProductGroup.id))).scalar_one()
    updated_at = db.execute(select(func.max(ProductGroup.updated_at))).scalar_one()
    upload_updated_at = db.execute(select(func.max(ProductGroup.upload_updated_at))).scalar_one()
    manual_updated_at = db.execute(select(func.max(ProductGroup.manual_updated_at))).scalar_one()
    return {
        "total_count": int(total_count or 0),
        "updated_at": updated_at.isoformat() if updated_at else None,
        "upload_updated_at": upload_updated_at.isoformat() if upload_updated_at else None,
        "manual_updated_at": manual_updated_at.isoformat() if manual_updated_at else None,
        "required_columns": MAPPING_TEMPLATE_COLUMNS,
        "optional_columns": OPTIONAL_MAPPING_COLUMNS,
    }


def _record_overseas_sku_locales(record: dict) -> list[dict]:
    return [
        loc
        for loc in (record.get("sku_locales") or [])
        if str(loc.get("country_code") or "").strip().upper() != "KR" and str(loc.get("sku") or "").strip()
    ]


def _record_kr_sku_from_row(record: dict) -> str:
    anchor = _normalize_mapping_sku(record.get("anchor_kr_sku"))
    if anchor:
        return anchor
    for loc in record.get("sku_locales") or []:
        if str(loc.get("country_code") or "").strip().upper() != "KR":
            continue
        sku = _normalize_mapping_sku(loc.get("sku"))
        if sku:
            return sku
    return ""


def _format_skipped_overseas_warning(overseas_sku: str, kr_sku: str = "") -> str:
    overseas = str(overseas_sku or "").strip() or "(코드 없음)"
    kr = str(kr_sku or "").strip()
    if kr:
        return (
            f"해외 상품 (상품코드 : {overseas})는 매핑되어 있는 한국 상품코드({kr})가 DB에 존재하지 않습니다.\n"
            "해당 상품이 실제로 한국에 존재하는 상품인지 확인해 주시고, "
            "존재한다면 해당 한국 상품을 먼저 등록해 주세요."
        )
    return (
        f"해외 상품 (상품코드 : {overseas})는 매핑되어 있는 한국 상품이 DB에 존재하지 않습니다.\n"
        "해당 상품이 실제로 한국에 존재하는 상품인지 확인해 주시고, "
        "존재한다면 해당 한국 상품을 먼저 등록해 주세요."
    )


def _collect_skipped_overseas_warnings(record: dict, kr_sku: str, seen: set[str], out: list[str]) -> None:
    for loc in _record_overseas_sku_locales(record):
        sku = str(loc.get("sku") or "").strip()
        if not sku or sku in seen:
            continue
        seen.add(sku)
        out.append(_format_skipped_overseas_warning(sku, kr_sku))


def merge_item_sku_mappings(db: Session, upload_files: list[UploadFile]) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")
    if not upload_files:
        raise HTTPException(status_code=400, detail="업로드할 SKU 파일이 필요합니다.")

    records = _read_merge_records(upload_files)
    uploaded_at = datetime.now(timezone.utc)
    groups = _load_product_groups(db)
    group_index: dict[uuid.UUID, ProductGroup] = {group.id: group for group in groups}
    sku_index: dict[tuple[str, str, str], ProductGroup] = {}
    name_index: dict[tuple[str, str], set[uuid.UUID]] = {}
    for group in groups:
        _index_group(group, sku_index, name_index, group_index)

    kr_tw_core_index = _build_kr_tw_core_index(groups)

    touched_group_ids: set[uuid.UUID] = set()
    skipped_overseas_warnings: list[str] = []
    skipped_overseas_seen: set[str] = set()
    try:
        for record in records:
            overseas_locales = _record_overseas_sku_locales(record)
            is_overseas_row = bool(overseas_locales)
            matched_groups = _find_matching_groups(
                record, sku_index, name_index, group_index, kr_tw_core_index
            )

            if is_overseas_row:
                kr_code = _record_kr_sku_from_row(record)
                if kr_code:
                    kr_hit = _lookup_group_by_country_sku_variants(sku_index, "KR", kr_code)
                    if kr_hit is None:
                        _collect_skipped_overseas_warnings(
                            record, kr_code, skipped_overseas_seen, skipped_overseas_warnings
                        )
                        continue
                    if not matched_groups:
                        matched_groups = [kr_hit]
                if not matched_groups:
                    # 해외 행은 DB에 한국 상품이 있을 때만 매핑. 신규 item 생성하지 않음.
                    _collect_skipped_overseas_warnings(
                        record, kr_code, skipped_overseas_seen, skipped_overseas_warnings
                    )
                    continue
                target = _merge_groups(
                    db=db,
                    groups=matched_groups,
                    preferred_kr_name=str(record["kr_name"] or ""),
                    sku_index=sku_index,
                    name_index=name_index,
                    group_index=group_index,
                    row_label=str(record["row_label"]),
                )
            elif matched_groups:
                target = _merge_groups(
                    db=db,
                    groups=matched_groups,
                    preferred_kr_name=str(record["kr_name"] or ""),
                    sku_index=sku_index,
                    name_index=name_index,
                    group_index=group_index,
                    row_label=str(record["row_label"]),
                )
            else:
                new_brand = _normalize_brand_cell(record.get("brand"))
                if not new_brand:
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            f"{record['row_label']}: DB에 없는 신규 상품(item)을 추가할 때는 "
                            f"`brand` 또는 `브랜드` 열에 값이 필요합니다."
                        ),
                    )
                new_barcode = _normalize_barcode_cell(record.get(BARCODE_COLUMN_CANONICAL))
                new_category = _normalize_category_cell(record.get(CATEGORY_COLUMN_CANONICAL))
                new_segment = normalize_item_segment(record.get("segment"))
                new_mkt = normalize_mkt_priority_cell(record.get("mkt_priority"))
                target = ProductGroup(
                    id=uuid.uuid4(),
                    kr_name=str(record["kr_name"] or record["fallback_name"]),
                    brand=new_brand,
                    representative_code=_resolve_representative_code_for_mapping(
                        record.get("representative_code"),
                        str(record.get("anchor_kr_sku") or "").strip()
                        or next(
                            (
                                str(loc.get("sku") or "")
                                for loc in record.get("sku_locales", [])
                                if str(loc.get("country_code") or "").upper() == "KR"
                            ),
                            "",
                        ),
                    ),
                    barcode=new_barcode if new_barcode else None,
                    category=new_category,
                    segment=new_segment,
                    mkt_priority=new_mkt,
                    upload_updated_at=uploaded_at,
                    manual_updated_at=None,
                    updated_at=uploaded_at,
                )
                db.add(target)
                group_index[target.id] = target
            target = _apply_merge_record(
                db=db,
                target=target,
                record=record,
                uploaded_at=uploaded_at,
                sku_index=sku_index,
                name_index=name_index,
                group_index=group_index,
            )
            touched_group_ids.add(target.id)
            # 대만 코어 인덱스 갱신(같은 배치 내 후속 행용)
            kr_tw_core_index = _build_kr_tw_core_index(list(group_index.values()))

        db.commit()
    except HTTPException:
        db.rollback()
        raise
    except Exception:
        db.rollback()
        logger.exception("SKU 매핑 병합 처리 중 예외")
        raise

    summary = get_item_sku_mapping_summary(db)
    summary["processed_file_count"] = len(upload_files)
    summary["merged_item_count"] = len(touched_group_ids)
    summary["skipped_overseas_count"] = len(skipped_overseas_warnings)
    summary["skipped_overseas_warnings"] = skipped_overseas_warnings
    return summary


def clear_item_sku_mappings(db: Session) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")

    deleted_count = db.execute(select(func.count(ProductGroup.id))).scalar_one()
    try:
        db.execute(delete(ProductLocale))
        db.execute(delete(ProductGroup))
        db.commit()
    except Exception:
        db.rollback()
        raise
    return {"deleted_count": int(deleted_count or 0)}


def upsert_item_sku_mapping(db: Session, payload: dict[str, object]) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")

    record = _build_mapping_group(payload, "수기 입력")
    rows = _load_product_groups(db)
    manual_updated_at = datetime.now(timezone.utc)
    target = _find_group_by_kr_anchor(rows, record)
    _validate_group_conflicts(rows, record, exclude_group_id=target.id if target else None)
    try:
        if target is None:
            bc0 = _normalize_barcode_cell(record.get(BARCODE_COLUMN_CANONICAL))
            cat0 = _normalize_category_cell(record.get(CATEGORY_COLUMN_CANONICAL))
            seg0 = normalize_item_segment(record.get("segment"))
            mkt0 = normalize_mkt_priority_cell(record.get("mkt_priority"))
            target = ProductGroup(
                kr_name=record["kr_name"],
                brand=record["brand"],
                representative_code=record.get("representative_code")
                or next(
                    (loc["sku"] for loc in record["locales"] if loc["country_code"] == "KR"),
                    None,
                ),
                barcode=bc0 if bc0 else None,
                category=cat0,
                segment=seg0,
                mkt_priority=mkt0,
                upload_updated_at=None,
                manual_updated_at=manual_updated_at,
                updated_at=manual_updated_at,
            )
            target.locales = [
                ProductLocale(
                    country_code=locale["country_code"],
                    sku_type=locale.get("sku_type"),
                    name=locale["name"],
                    sku=locale["sku"],
                    updated_at=manual_updated_at,
                )
                for locale in record["locales"]
            ]
            db.add(target)
        else:
            _merge_manual_mapping_locales(target, record, manual_updated_at)
        db.commit()
        db.refresh(target)
    except Exception:
        db.rollback()
        raise

    return _mapping_item_payload(target)


def patch_item_sku_mapping_segment(db: Session, group_id: str, discontinued: bool) -> dict:
    """한국 상품 그룹(item) 단종 여부만 수동 변경. DB에는 단종 문자열만 저장."""
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")

    raw = str(group_id or "").strip()
    try:
        gid = uuid.UUID(raw)
    except ValueError:
        raise HTTPException(status_code=400, detail="유효하지 않은 group_id입니다.") from None

    group = db.get(ProductGroup, gid)
    if group is None:
        raise HTTPException(status_code=404, detail="상품 그룹을 찾을 수 없습니다.")

    now = datetime.now(timezone.utc)
    if discontinued:
        group.segment = "단종"
    elif str(group.segment or "").strip() == "단종":
        group.segment = None
    group.manual_updated_at = now
    group.updated_at = now

    try:
        db.commit()
        fresh = db.execute(
            select(ProductGroup).where(ProductGroup.id == gid).options(selectinload(ProductGroup.locales))
        ).scalar_one()
    except Exception:
        db.rollback()
        raise

    return _mapping_item_payload(fresh)


def patch_item_sku_mapping_item(db: Session, payload: dict[str, object]) -> dict:
    """한국 상품 그룹: 상품코드·바코드·브랜드·상품명·마케팅 우선순위·구분 일괄 수정."""
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
        raise HTTPException(status_code=404, detail="상품 그룹을 찾을 수 없습니다.")

    kr_locale = next(
        (loc for loc in group.locales if str(loc.country_code or "").strip().upper() == "KR"),
        None,
    )
    if kr_locale is None:
        raise HTTPException(status_code=400, detail="한국 상품코드(로케일)가 없어 수정할 수 없습니다.")

    kr_sku_in = str(payload.get("kr_sku") or "").strip()
    kr_name_in = str(payload.get("kr_name") or "").strip()
    brand_in = str(payload.get("brand") or "").strip()

    if not kr_sku_in:
        raise HTTPException(status_code=400, detail="상품코드는 비울 수 없습니다.")
    if not kr_name_in:
        raise HTTPException(status_code=400, detail="상품명은 비울 수 없습니다.")
    if not brand_in:
        raise HTTPException(status_code=400, detail="브랜드는 비울 수 없습니다.")

    new_sku_norm = _normalize_mapping_sku(kr_sku_in)
    if not new_sku_norm:
        raise HTTPException(status_code=400, detail="상품코드가 비어 있습니다.")

    old_norm = _normalize_mapping_sku(kr_locale.sku)
    if new_sku_norm != old_norm:
        others = db.execute(
            select(ProductLocale).where(
                ProductLocale.country_code == "KR",
                ProductLocale.item_id != gid,
            )
        ).scalars().all()
        for loc in others:
            if _normalize_mapping_sku(loc.sku) == new_sku_norm:
                raise HTTPException(
                    status_code=400,
                    detail=f"다른 상품이 이미 사용 중인 한국 상품코드입니다: {kr_sku_in}",
                )
        kr_locale.sku = kr_sku_in

    group.kr_name = kr_name_in
    kr_locale.name = kr_name_in
    group.brand = _normalize_brand_cell(brand_in)

    if "barcode" in payload:
        bc = _normalize_barcode_cell(payload.get("barcode"))
        group.barcode = bc if bc else None

    if "category" in payload:
        group.category = _normalize_category_cell(payload.get("category"))

    if "mkt_priority" in payload:
        mkt_s = str(payload.get("mkt_priority") or "").strip()
        group.mkt_priority = normalize_mkt_priority_cell(mkt_s) if mkt_s else None

    if "segment" in payload:
        seg_s = str(payload.get("segment") or "").strip()
        group.segment = normalize_item_segment(seg_s) if seg_s else None

    now = datetime.now(timezone.utc)
    group.manual_updated_at = now
    group.updated_at = now
    kr_locale.updated_at = now

    try:
        db.commit()
        fresh = db.execute(
            select(ProductGroup).where(ProductGroup.id == gid).options(selectinload(ProductGroup.locales))
        ).scalar_one()
    except Exception:
        db.rollback()
        raise

    return _mapping_item_payload(fresh)


def list_item_sku_mappings(db: Session, query: str | None = None, limit: int = 100) -> list[dict]:
    rows = _load_product_groups(db)
    needle = str(query or "").strip().casefold()
    cap = max(1, limit)

    if not needle:
        # 검색 전: 국가 수 → 채워진 필드 수 → 단종 후순위 — limit slice 전 정렬
        ordered = sorted(
            rows,
            key=lambda g: (
                -len(g.locales or []),
                -_filled_product_search_detail_count(g),
                1 if _mapping_item_discontinued(g) else 0,
                str(g.kr_name or "").casefold(),
            ),
        )
        return [_mapping_item_payload(g) for g in ordered[:cap]]

    items: list[dict] = []
    for group in rows:
        haystack_parts = [
            group.kr_name,
            str(group.barcode or ""),
            str(group.category or ""),
            format_item_segment_display(normalize_item_segment(group.segment)),
            str(group.segment or ""),
            str(group.mkt_priority or ""),
        ]
        for locale in group.locales:
            haystack_parts.extend([locale.country_code, locale.name or "", locale.sku])
        haystack = " ".join(haystack_parts).casefold()
        if needle not in haystack:
            continue
        items.append(_mapping_item_payload(group))
        if len(items) >= cap:
            break
    return items


def build_item_sku_lookup_maps(db: Session) -> dict:
    groups = _load_product_groups(db)
    by_country_sku: dict[str, dict[str, dict[str, str]]] = {code: {} for code in COUNTRY_CODES}
    by_country_name: dict[str, dict[str, dict[str, str]]] = {code: {} for code in COUNTRY_CODES}

    for group in groups:
        kr_locale = next((locale for locale in group.locales if locale.country_code == "KR"), None)
        seg_norm = normalize_item_segment(group.segment)
        mkt_norm = normalize_mkt_priority_cell(group.mkt_priority)
        payload = {
            "group_id": str(group.id),
            "kr_name": kr_locale.name if kr_locale else group.kr_name,
            "kr_sku": kr_locale.sku if kr_locale else "",
            "segment": seg_norm,
            "segment_display": format_item_segment_display(seg_norm),
            "mkt_priority": mkt_norm,
        }
        for locale in group.locales:
            by_country_sku[locale.country_code][locale.sku] = payload
            if locale.name:
                nk = _normalize_mapping_name_key(locale.name)
                by_country_name[locale.country_code][nk] = payload

    return {"by_country_sku": by_country_sku, "by_country_name": by_country_name}


def resolve_item_sku_mapping(
    lookups: dict,
    country_code: str,
    raw_item_code: object,
    description: object = "",
) -> dict[str, str]:
    normalized_country = str(country_code or "").strip().upper() or "KR"
    normalized_sku = _normalize_mapping_sku(raw_item_code)
    country_rows = ((lookups or {}).get("by_country_sku") or {}).get(normalized_country, {})
    matched = country_rows.get(normalized_sku) if normalized_sku else None
    if not matched:
        matched = ((lookups or {}).get("by_country_name") or {}).get(normalized_country, {}).get(
            _normalize_mapping_name_key(description)
        )
    if not matched:
        return {
            "mapped_kr_name": "",
            "mapped_kr_sku": "",
            "mapped_segment": "",
            "mapped_segment_display": "",
            "mapped_mkt_priority": "",
        }

    seg_raw = matched.get("segment")
    seg_disp = matched.get("segment_display")
    mkt_raw = matched.get("mkt_priority")
    return {
        "mapped_kr_name": str(matched.get("kr_name") or ""),
        "mapped_kr_sku": str(matched.get("kr_sku") or ""),
        "mapped_segment": "" if seg_raw is None else str(seg_raw).strip(),
        "mapped_segment_display": "" if seg_disp is None else str(seg_disp),
        "mapped_mkt_priority": "" if mkt_raw is None else str(mkt_raw).strip(),
    }


def find_item_by_any_country_sku(db: Session, raw_sku: object) -> dict[str, str | bool]:
    """발주 등록용: 어느 국가 SKU든 매칭되면 브랜드·한국 기준 상품명 반환."""
    needle = _normalize_mapping_sku(raw_sku)
    if not needle:
        return {
            "matched": False,
            "brand": "",
            "product_name": "",
            "kr_sku": "",
            "message": "상품코드를 입력하세요.",
        }
    groups = _load_product_groups(db)
    for group in groups:
        for loc in group.locales:
            if _normalize_mapping_sku(loc.sku) == needle:
                kr_loc = next((x for x in group.locales if x.country_code == "KR"), None)
                name = (kr_loc.name if kr_loc and kr_loc.name else None) or group.kr_name or ""
                kr_sku_val = str(kr_loc.sku if kr_loc and kr_loc.sku else "").strip() or needle
                return {
                    "matched": True,
                    "brand": str(group.brand or "").strip(),
                    "product_name": str(name).strip(),
                    "kr_sku": kr_sku_val,
                    "message": "",
                }
    return {
        "matched": False,
        "brand": "",
        "product_name": "",
        "kr_sku": "",
        "message": "등록된 상품코드가 없습니다. SKU 관리 탭에서 상품코드를 먼저 등록하세요.",
    }


__all__ = [
    "build_item_sku_lookup_maps",
    "clear_item_sku_mappings",
    "find_item_by_any_country_sku",
    "get_item_sku_mapping_summary",
    "list_item_sku_mappings",
    "merge_item_sku_mappings",
    "patch_item_sku_mapping_item",
    "patch_item_sku_mapping_segment",
    "resolve_item_sku_mapping",
    "upsert_item_sku_mapping",
]
