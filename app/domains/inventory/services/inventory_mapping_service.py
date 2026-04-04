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

from app.domains.inventory.models import ProductGroup, ProductLocale
from app.domains.inventory.services.inventory_aggregate_service import _normalize_item_code
from app.shared.config import get_runtime_settings

logger = logging.getLogger(__name__)

COUNTRY_FIELD_SPECS = [
    ("KR", "kr_name", "kr_sku"),
    ("US", "us_name", "us_sku"),
    ("TW", "tw_name", "tw_sku"),
    ("VN", "vn_name", "vn_sku"),
    ("SG", "sg_name", "sg_sku"),
    ("AU", "au_name", "au_sku"),
    ("UK", "uk_name", "uk_sku"),
    ("AE", "ae_name", "ae_sku"),
]
COUNTRY_ORDER = {country_code: index for index, (country_code, _, _) in enumerate(COUNTRY_FIELD_SPECS)}
COUNTRY_CODES = [country_code for country_code, _, _ in COUNTRY_FIELD_SPECS]
MAPPING_TEMPLATE_COLUMNS = [column for _, name_key, sku_key in COUNTRY_FIELD_SPECS for column in (name_key, sku_key)]
MAX_MAPPING_FILE_SIZE_BYTES = 2 * 1024 * 1024
# AE/UK/AU 엑셀은 전용 name 열 대신 us_name 만 있는 경우가 많음
MAPPING_NAME_FALLBACK_TO_US_NAME = frozenset({"AU", "UK", "AE"})

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
    raw = str(value or "").strip()
    if not raw:
        return ""
    return _normalize_item_code(raw)


def _normalize_mapping_header(value: object) -> str:
    raw = str(value or "")
    normalized = raw.replace("\ufeff", "").replace("\u200b", "").strip().casefold()
    normalized = re.sub(r"\s+", "", normalized)
    normalized = re.sub(r"[_\-()]", "", normalized)
    return normalized


def _lookup_group_by_country_sku_variants(
    sku_index: dict[tuple[str, str], ProductGroup],
    country_code: str,
    raw_sku: object,
) -> ProductGroup | None:
    sku = _normalize_mapping_sku(raw_sku)
    if not sku:
        return None
    cc = str(country_code or "").strip().upper()
    cur: str | None = sku
    seen: set[str] = set()
    while cur and cur not in seen:
        seen.add(cur)
        hit = sku_index.get(_index_key(cc, cur))
        if hit is not None:
            return hit
        nxt = _sku_strip_one_variant_suffix(cur)
        if not nxt or nxt == cur:
            break
        cur = nxt
    return None


def _tw_hk_five_digit_core_from_tw_sku(tw_sku: object) -> str | None:
    """HK06019-02 → 앞쪽 국가코드(영문) 뒤의 5자리 숫자 = 한국 SKU 코어."""
    s = _normalize_mapping_sku(tw_sku)
    if not s:
        return None
    m = re.match(r"^[A-Za-z]+(\d{5})(?:-.+)?$", s)
    return m.group(1) if m else None


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


def _read_mapping_frame(upload_file: UploadFile) -> pd.DataFrame:
    """허용 컬럼만 사용하고, 나머지 열은 무시한다."""
    raw, _ = _read_mapping_bytes(upload_file)
    filename = str(upload_file.filename or "mapping.xlsx")
    try:
        df = pd.read_excel(io.BytesIO(raw), sheet_name=0, dtype=object)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"{filename}: 엑셀 파싱 실패 ({exc})") from exc

    header_key_to_canonical: dict[str, str] = {}
    for canonical in MAPPING_TEMPLATE_COLUMNS:
        key = _normalize_mapping_header(canonical)
        if key in header_key_to_canonical:
            raise HTTPException(status_code=500, detail="매핑 템플릿 컬럼 정의가 서로 충돌합니다.")
        header_key_to_canonical[key] = canonical

    series_by_canonical: dict[str, pd.Series] = {}
    matched_raw: dict[str, str] = {}

    for col in df.columns:
        raw_label = str(col).strip()
        norm = _normalize_mapping_header(raw_label)
        if norm not in header_key_to_canonical:
            continue
        canonical = header_key_to_canonical[norm]
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

    if not series_by_canonical:
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
    for canonical in MAPPING_TEMPLATE_COLUMNS:
        if canonical in series_by_canonical:
            aligned[canonical] = series_by_canonical[canonical]
        else:
            aligned[canonical] = pd.Series([pd.NA] * len(df), index=df.index, dtype=object)
    return pd.DataFrame(aligned)


def _mapping_item_payload(group: ProductGroup) -> dict:
    locales = sorted(group.locales, key=lambda locale: COUNTRY_ORDER.get(locale.country_code, 999))
    return {
        "group_id": str(group.id),
        "kr_name": group.kr_name,
        "locales": [
            {
                "country_code": locale.country_code,
                "name": locale.name,
                "sku": locale.sku,
            }
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
    locales: list[dict[str, str]] = []
    for country_code, name_key, sku_key in COUNTRY_FIELD_SPECS:
        name = _resolve_row_localized_name(source, country_code, name_key)
        sku = _normalize_mapping_sku(source.get(sku_key))
        if not name and not sku:
            continue
        if not name or not sku:
            raise HTTPException(
                status_code=400,
                detail=f"{row_label}: {country_code} 상품명과 SKU는 함께 입력해야 합니다.",
            )
        locales.append({"country_code": country_code, "name": name, "sku": sku})
    if not locales:
        raise HTTPException(status_code=400, detail=f"{row_label}: 최소 1개 이상의 국가 매핑이 필요합니다.")
    kr_locale = next((locale for locale in locales if locale["country_code"] == "KR"), None)
    if kr_locale is None:
        raise HTTPException(status_code=400, detail=f"{row_label}: KR 상품명과 KR SKU는 필수입니다.")
    return {
        "kr_name": kr_locale["name"],
        "locales": locales,
    }


def _find_group_by_kr_anchor(groups: list[ProductGroup], payload: dict) -> ProductGroup | None:
    kr_locale = next((locale for locale in payload["locales"] if locale["country_code"] == "KR"), None)
    if kr_locale is None:
        return None
    target_name_key = _normalize_mapping_name_key(kr_locale["name"])
    target_sku = kr_locale["sku"]
    for group in groups:
        for locale in group.locales:
            if locale.country_code != "KR":
                continue
            if locale.sku == target_sku or _normalize_mapping_name_key(locale.name) == target_name_key:
                return group
    return None


def _validate_group_conflicts(groups: list[ProductGroup], payload: dict, exclude_group_id: uuid.UUID | None = None) -> None:
    seen_keys: set[tuple[str, str]] = set()
    for locale in payload["locales"]:
        key = (locale["country_code"], locale["sku"])
        if key in seen_keys:
            raise HTTPException(
                status_code=400,
                detail=f"{locale['country_code']} SKU `{locale['sku']}` 값이 중복됩니다.",
            )
        seen_keys.add(key)

    for group in groups:
        if exclude_group_id is not None and group.id == exclude_group_id:
            continue
        for locale in group.locales:
            for incoming in payload["locales"]:
                if locale.country_code == incoming["country_code"] and locale.sku == incoming["sku"]:
                    raise HTTPException(
                        status_code=400,
                        detail=f"{incoming['country_code']} SKU `{incoming['sku']}` 는 이미 다른 상품 그룹에 등록되어 있습니다.",
                    )


def _upsert_group_locales(target: ProductGroup, payload: dict, now: datetime) -> None:
    target.kr_name = payload["kr_name"]
    target.manual_updated_at = now
    target.locales.clear()
    target.locales.extend(
        [
            ProductLocale(
                country_code=locale["country_code"],
                name=locale["name"],
                sku=locale["sku"],
                updated_at=now,
            )
            for locale in payload["locales"]
        ]
    )


def _parse_merge_record(source: dict[str, object], row_label: str) -> dict | None:
    sku_locales: list[dict[str, str | None]] = []
    named_locales: list[dict[str, str]] = []

    for country_code, name_key, sku_key in COUNTRY_FIELD_SPECS:
        name = _resolve_row_localized_name(source, country_code, name_key)
        sku = _normalize_mapping_sku(source.get(sku_key))
        if not name and not sku:
            continue
        if sku:
            clean_name = None
            if name and not _is_placeholder_mapping_name(name):
                clean_name = name
            sku_locales.append({"country_code": country_code, "name": clean_name, "sku": sku})
        elif name and not _is_placeholder_mapping_name(name):
            named_locales.append({"country_code": country_code, "name": name})

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
    match_us_name = _normalize_mapping_name(source.get("us_name"))
    if _is_placeholder_mapping_name(match_us_name):
        match_us_name = ""

    return {
        "row_label": row_label,
        "kr_name": anchor_kr_name,
        "sku_locales": sku_locales,
        "named_locales": named_locales,
        "fallback_name": first_name or first_sku or "미분류 상품",
        "anchor_kr_sku": anchor_kr_sku,
        "anchor_kr_name": anchor_kr_name,
        "match_us_name": match_us_name,
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


def _locale_tuple_key(country_code: str, sku: str) -> tuple[str, str]:
    return (str(country_code or "").strip().upper(), str(sku or "").strip())


def _locales_map_by_country_sku(group: ProductGroup) -> dict[tuple[str, str], ProductLocale]:
    return {_locale_tuple_key(loc.country_code, loc.sku): loc for loc in group.locales}


def _remove_group_from_indexes(
    group: ProductGroup,
    sku_index: dict[tuple[str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
) -> None:
    for locale in group.locales:
        sku_key = _index_key(locale.country_code, locale.sku)
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
    sku_index: dict[tuple[str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
) -> None:
    group_index[group.id] = group
    for locale in group.locales:
        sku_index[_index_key(locale.country_code, locale.sku)] = group
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
    priority = ("KR", "TW", "US", "VN", "AE", "AU", "UK", "SG")
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
    sku_index: dict[tuple[str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
    kr_tw_core_index: dict[str, set[uuid.UUID]],
) -> list[ProductGroup]:
    matched_ids: set[uuid.UUID] = set()
    row_label = str(record["row_label"])
    # 5)·6)에서 사용: 이 국가는 이미 SKU/코어 등으로 item 이 잠겼으면 상품명으로 다시 묶지 않음
    sku_resolved_cc: set[str] = set()

    # 1) 대만: HK06019 / HK06019-02 → 동일 5자리 한국 코어 상품
    #    코어 추출 불가 시 → 엑셀 us_name 과 DB US 로케일 상품명으로 한국 item 찾기
    tw_cores: list[str] = []
    has_tw_locale = False
    for loc in record["sku_locales"]:
        if str(loc.get("country_code") or "").strip().upper() != "TW":
            continue
        has_tw_locale = True
        core = _tw_hk_five_digit_core_from_tw_sku(loc.get("sku"))
        if core:
            tw_cores.append(core)
    if tw_cores:
        locked_tw: set[uuid.UUID] = set()
        for core in tw_cores:
            hits = kr_tw_core_index.get(core, set())
            if len(hits) > 1:
                raise HTTPException(
                    status_code=400,
                    detail=(
                        f"{row_label}: 대만 품번 `{core}` 에 해당하는 한국 SKU 를 가진 상품이 DB 에 여러 개입니다. "
                        "한국 마스터에서 해당 5자리를 유일하게 맞춘 뒤 다시 업로드하세요."
                    ),
                )
            if len(hits) == 1:
                locked_tw.add(next(iter(hits)))
        if len(locked_tw) > 1:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{row_label}: 같은 행의 대만 SKU 들이 서로 다른 한국 상품(5자리 코어)을 가리킵니다. "
                    "한 행에는 동일 상품의 대만 코드만 넣어 주세요."
                ),
            )
        if len(locked_tw) == 1:
            matched_ids.add(next(iter(locked_tw)))
            sku_resolved_cc.add("TW")
            sku_resolved_cc.add("KR")
    elif has_tw_locale:
        us_for_tw = _normalize_mapping_name(record.get("match_us_name"))
        if not us_for_tw or _is_placeholder_mapping_name(us_for_tw):
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{row_label}: 대만 `tw_sku` 에서 5자리 품번 규칙(예: HK+5자리)을 찾지 못했습니다. "
                    "같은 행에 `us_name`(미국 상품명)을 넣어 기존 상품과 연결하거나, 품번 형식을 맞춰 주세요."
                ),
            )
        tw_us_candidates = name_index.get(_index_key("US", _normalize_mapping_name_key(us_for_tw)), set())
        if len(tw_us_candidates) > 1:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{row_label}: 대만 행의 `us_name` `{us_for_tw}` 과 일치하는 US 상품명을 가진 기존 상품이 여러 개입니다. "
                    "US 명을 유일하게 맞추거나 `kr_sku` 로 먼저 연결해 주세요."
                ),
            )
        if len(tw_us_candidates) == 1:
            _add_matched_group_id(
                matched_ids,
                next(iter(tw_us_candidates)),
                row_label,
                f"대만 SKU 에 코어가 없어 `us_name`={us_for_tw!r} 로 찾은 상품과",
                group_index,
                record,
            )
            sku_resolved_cc.add("TW")
            sku_resolved_cc.add("KR")

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
    for locale in record["sku_locales"]:
        cc = str(locale["country_code"] or "").strip().upper()
        sku = str(locale["sku"] or "")
        matched = _lookup_group_by_country_sku_variants(sku_index, cc, sku)
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

    # 6) 국가별 상품명
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
    sku_index: dict[tuple[str, str], ProductGroup],
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
            lk = _locale_tuple_key(loc.country_code, loc.sku)
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
    sku_index: dict[tuple[str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
) -> ProductGroup:
    _remove_group_from_indexes(target, sku_index, name_index, group_index)

    locales_map = _locales_map_by_country_sku(target)
    for locale in record["sku_locales"]:
        country_code = str(locale["country_code"])
        sku = str(locale["sku"] or "")
        name = _normalize_mapping_name(locale.get("name"))
        lk = _locale_tuple_key(country_code, sku)
        conflict_group = sku_index.get(_index_key(country_code, sku))
        if conflict_group is not None and conflict_group.id != target.id:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{record['row_label']}: {country_code} SKU `{sku}` 는 이미 다른 상품에 연결되어 있습니다. "
                    "충돌이 있어 전체 업로드를 중단합니다."
                ),
            )
        existing = locales_map.get(lk)
        if existing is None:
            new_locale = ProductLocale(
                country_code=country_code,
                name=name or None,
                sku=sku,
                updated_at=uploaded_at,
            )
            target.locales.append(new_locale)
            locales_map[lk] = new_locale
            continue
        if name:
            existing.name = name
        existing.updated_at = uploaded_at

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
    target.upload_updated_at = uploaded_at
    target.updated_at = uploaded_at

    _index_group(target, sku_index, name_index, group_index)
    return target


def get_product_sku_mapping_summary(db: Session) -> dict:
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
    }


def merge_product_sku_mappings(db: Session, upload_files: list[UploadFile]) -> dict:
    settings = get_runtime_settings()
    if not settings.database_enabled:
        raise HTTPException(status_code=500, detail="DATABASE_URL이 설정되지 않았습니다.")
    if not upload_files:
        raise HTTPException(status_code=400, detail="업로드할 SKU 파일이 필요합니다.")

    records = _read_merge_records(upload_files)
    uploaded_at = datetime.now(timezone.utc)
    groups = _load_product_groups(db)
    group_index: dict[uuid.UUID, ProductGroup] = {group.id: group for group in groups}
    sku_index: dict[tuple[str, str], ProductGroup] = {}
    name_index: dict[tuple[str, str], set[uuid.UUID]] = {}
    for group in groups:
        _index_group(group, sku_index, name_index, group_index)

    kr_tw_core_index = _build_kr_tw_core_index(groups)

    touched_group_ids: set[uuid.UUID] = set()
    try:
        for record in records:
            matched_groups = _find_matching_groups(
                record, sku_index, name_index, group_index, kr_tw_core_index
            )
            if matched_groups:
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
                target = ProductGroup(
                    id=uuid.uuid4(),
                    kr_name=str(record["kr_name"] or record["fallback_name"]),
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

    summary = get_product_sku_mapping_summary(db)
    summary["processed_file_count"] = len(upload_files)
    summary["merged_item_count"] = len(touched_group_ids)
    return summary


def clear_product_sku_mappings(db: Session) -> dict:
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


def upsert_product_sku_mapping(db: Session, payload: dict[str, object]) -> dict:
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
            target = ProductGroup(
                kr_name=record["kr_name"],
                upload_updated_at=None,
                manual_updated_at=manual_updated_at,
                updated_at=manual_updated_at,
            )
            target.locales = [
                ProductLocale(
                    country_code=locale["country_code"],
                    name=locale["name"],
                    sku=locale["sku"],
                    updated_at=manual_updated_at,
                )
                for locale in record["locales"]
            ]
            db.add(target)
        else:
            _upsert_group_locales(target, record, manual_updated_at)
        db.commit()
        db.refresh(target)
    except Exception:
        db.rollback()
        raise

    return _mapping_item_payload(target)


def list_product_sku_mappings(db: Session, query: str | None = None, limit: int = 100) -> list[dict]:
    rows = _load_product_groups(db)
    needle = str(query or "").strip().casefold()
    items: list[dict] = []
    for group in rows:
        item = _mapping_item_payload(group)
        if needle:
            haystack_parts = [group.kr_name]
            for locale in group.locales:
                haystack_parts.extend([locale.country_code, locale.name or "", locale.sku])
            haystack = " ".join(haystack_parts).casefold()
            if needle not in haystack:
                continue
        items.append(item)
        if len(items) >= max(1, limit):
            break
    return items


def build_product_sku_mapping_lookups(db: Session) -> dict:
    groups = _load_product_groups(db)
    by_country_sku: dict[str, dict[str, dict[str, str]]] = {code: {} for code in COUNTRY_CODES}
    by_country_name: dict[str, dict[str, dict[str, str]]] = {code: {} for code in COUNTRY_CODES}

    for group in groups:
        kr_locale = next((locale for locale in group.locales if locale.country_code == "KR"), None)
        payload = {
            "group_id": str(group.id),
            "kr_name": kr_locale.name if kr_locale else group.kr_name,
            "kr_sku": kr_locale.sku if kr_locale else "",
        }
        for locale in group.locales:
            by_country_sku[locale.country_code][locale.sku] = payload
            if locale.name:
                nk = _normalize_mapping_name_key(locale.name)
                by_country_name[locale.country_code][nk] = payload

    return {"by_country_sku": by_country_sku, "by_country_name": by_country_name}


def resolve_product_sku_mapping(
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
        return {"mapped_kr_name": "", "mapped_kr_sku": ""}

    return {
        "mapped_kr_name": str(matched.get("kr_name") or ""),
        "mapped_kr_sku": str(matched.get("kr_sku") or ""),
    }
