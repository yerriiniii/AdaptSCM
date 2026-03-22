import pandas as pd

from app.shared.config.sheets import REQUIRED_SHEETS


def validate_input_frames(
    sales_df: pd.DataFrame, inventory_df: pd.DataFrame, policy_df: pd.DataFrame
) -> list[str]:
    """Validate required columns and basic value constraints."""
    errors = []
    data_map = {"sales": sales_df, "inventory": inventory_df, "sku_policy": policy_df}

    for sheet_name, required_cols in REQUIRED_SHEETS.items():
        df = data_map[sheet_name]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            errors.append(f"[{sheet_name}] 누락 컬럼: {missing}")

    if errors:
        return errors

    sales_df["date"] = pd.to_datetime(sales_df["date"], errors="coerce")
    inventory_df["date"] = pd.to_datetime(inventory_df["date"], errors="coerce")

    if sales_df["date"].isna().any():
        errors.append("[sales] date 파싱 실패 행이 있습니다.")
    if inventory_df["date"].isna().any():
        errors.append("[inventory] date 파싱 실패 행이 있습니다.")

    numeric_checks = [
        ("sales", "qty_sold", sales_df),
        ("inventory", "on_hand", inventory_df),
        ("inventory", "inbound_qty", inventory_df),
        ("sku_policy", "lead_time_days", policy_df),
        ("sku_policy", "moq", policy_df),
        ("sku_policy", "pack_size", policy_df),
        ("sku_policy", "safety_stock", policy_df),
    ]

    for sheet_name, col, df in numeric_checks:
        if (pd.to_numeric(df[col], errors="coerce").isna()).any():
            errors.append(f"[{sheet_name}] {col} 숫자 변환 실패 값이 있습니다.")

    if (sales_df["qty_sold"] < 0).any():
        errors.append("[sales] qty_sold는 음수일 수 없습니다.")
    if (inventory_df["on_hand"] < 0).any():
        errors.append("[inventory] on_hand는 음수일 수 없습니다.")
    if (inventory_df["inbound_qty"] < 0).any():
        errors.append("[inventory] inbound_qty는 음수일 수 없습니다.")

    return errors

