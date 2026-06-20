import re


def normalize_item_code(value: object) -> str:
    raw = str(value or "").strip()
    if not raw:
        return ""
    raw = re.sub(r"\.0$", "", raw)
    if re.fullmatch(r"\d+", raw):
        return raw.zfill(5) if len(raw) <= 5 else raw
    return raw.upper()


__all__ = ["normalize_item_code"]
