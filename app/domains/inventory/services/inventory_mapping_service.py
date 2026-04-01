import io
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


def _normalize_mapping_name(value: object) -> str:
    raw = str(value or "").strip()
    return re.sub(r"\s+", " ", raw)


def _normalize_mapping_name_key(value: object) -> str:
    return _normalize_mapping_name(value).casefold()


def _normalize_mapping_sku(value: object) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    return _normalize_item_code(raw)


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
    raw, _ = _read_mapping_bytes(upload_file)
    filename = str(upload_file.filename or "mapping.xlsx")
    try:
        df = pd.read_excel(io.BytesIO(raw), sheet_name=0, dtype=object)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"{filename}: 엑셀 파싱 실패 ({exc})") from exc

    actual_columns = [str(column).strip() for column in df.columns]
    if len(set(actual_columns)) != len(actual_columns):
        raise HTTPException(status_code=400, detail=f"{filename}: 헤더가 중복됩니다.")
    unknown_columns = [column for column in actual_columns if column not in MAPPING_TEMPLATE_COLUMNS]
    if unknown_columns:
        raise HTTPException(
            status_code=400,
            detail=(
                f"{filename}: 허용되지 않는 헤더가 있습니다. "
                f"허용 헤더: {', '.join(MAPPING_TEMPLATE_COLUMNS)} | "
                f"입력 헤더: {', '.join(actual_columns)}"
            ),
        )
    if df.empty:
        raise HTTPException(status_code=400, detail=f"{filename}: 최소 1개 이상의 매핑 행이 필요합니다.")
    df.columns = actual_columns
    return df


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
        name = _normalize_mapping_name(source.get(name_key))
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
        name = _normalize_mapping_name(source.get(name_key))
        sku = _normalize_mapping_sku(source.get(sku_key))
        if not name and not sku:
            continue
        if sku:
            sku_locales.append({"country_code": country_code, "name": name or None, "sku": sku})
        elif name:
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
    first_sku = next((str(locale.get("sku") or "").strip() for locale in sku_locales if str(locale.get("sku") or "").strip()), "")
    return {
        "row_label": row_label,
        "kr_name": _normalize_mapping_name(source.get("kr_name")),
        "sku_locales": sku_locales,
        "named_locales": named_locales,
        "fallback_name": first_name or first_sku or "미분류 상품",
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


def _get_group_locales_by_country(group: ProductGroup) -> dict[str, ProductLocale]:
    return {locale.country_code: locale for locale in group.locales}


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


def _find_matching_groups(
    record: dict,
    sku_index: dict[tuple[str, str], ProductGroup],
    name_index: dict[tuple[str, str], set[uuid.UUID]],
    group_index: dict[uuid.UUID, ProductGroup],
) -> list[ProductGroup]:
    matched_ids: set[uuid.UUID] = set()

    for locale in record["sku_locales"]:
        matched = sku_index.get(_index_key(locale["country_code"], str(locale["sku"] or "")))
        if matched is not None:
            matched_ids.add(matched.id)

    for locale in [*record["sku_locales"], *record["named_locales"]]:
        name = str(locale.get("name") or "").strip()
        if not name:
            continue
        candidate_ids = name_index.get(_index_key(locale["country_code"], _normalize_mapping_name_key(name)), set())
        if len(candidate_ids) <= 1:
            matched_ids.update(candidate_ids)
            continue
        if matched_ids and matched_ids.intersection(candidate_ids):
            continue
        raise HTTPException(
            status_code=400,
            detail=(
                f"{record['row_label']}: {locale['country_code']} 상품명 `{name}` 이 여러 상품과 연결되어 자동 병합할 수 없습니다."
            ),
        )

    return [group_index[group_id] for group_id in matched_ids if group_id in group_index]


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
    secondary_groups = [group for group in groups if group.id != primary.id]

    for group in [primary, *secondary_groups]:
        _remove_group_from_indexes(group, sku_index, name_index, group_index)

    primary_locales = _get_group_locales_by_country(primary)
    for secondary in secondary_groups:
        for locale in list(secondary.locales):
            existing = primary_locales.get(locale.country_code)
            if existing is not None:
                if existing.sku != locale.sku:
                    raise HTTPException(
                        status_code=400,
                        detail=(
                            f"{row_label}: 같은 국가 `{locale.country_code}` 에 서로 다른 SKU가 연결된 기존 상품이 있어 병합할 수 없습니다."
                        ),
                    )
                if locale.name and not existing.name:
                    existing.name = locale.name
                continue
            locale.product_group = primary
            primary_locales[locale.country_code] = locale
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

    locales_by_country = _get_group_locales_by_country(target)
    for locale in record["sku_locales"]:
        country_code = str(locale["country_code"])
        sku = str(locale["sku"] or "")
        name = _normalize_mapping_name(locale.get("name"))
        existing = locales_by_country.get(country_code)
        conflict_group = sku_index.get(_index_key(country_code, sku))
        if conflict_group is not None and conflict_group.id != target.id:
            raise HTTPException(
                status_code=400,
                detail=(
                    f"{record['row_label']}: {country_code} SKU `{sku}` 는 이미 다른 상품에 연결되어 있습니다. "
                    "충돌이 있어 전체 업로드를 중단합니다."
                ),
            )
        if existing is None:
            new_locale = ProductLocale(
                country_code=country_code,
                name=name or None,
                sku=sku,
                updated_at=uploaded_at,
            )
            target.locales.append(new_locale)
            locales_by_country[country_code] = new_locale
            continue
        existing.sku = sku
        if name:
            existing.name = name
        existing.updated_at = uploaded_at

    for locale in record["named_locales"]:
        existing = locales_by_country.get(str(locale["country_code"]))
        if existing is None:
            continue
        existing.name = _normalize_mapping_name(locale["name"])
        existing.updated_at = uploaded_at

    if record["kr_name"]:
        target.kr_name = record["kr_name"]
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

    touched_group_ids: set[uuid.UUID] = set()
    try:
        for record in records:
            matched_groups = _find_matching_groups(record, sku_index, name_index, group_index)
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

        db.commit()
    except Exception:
        db.rollback()
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
                by_country_name[locale.country_code][_normalize_mapping_name_key(locale.name)] = payload

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
