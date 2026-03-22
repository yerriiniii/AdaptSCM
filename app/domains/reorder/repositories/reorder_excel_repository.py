import io
from datetime import datetime, timedelta

import numpy as np
import pandas as pd


def create_dummy_data() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Generate sample data so users can test without real files."""
    np.random.seed(42)

    skus = ["SKU_001", "SKU_002", "SKU_003", "SKU_004"]
    channels = ["D2C", "OliveYoung", "MarketB"]
    sku_brand_map = {
        "SKU_001": "A_Brand",
        "SKU_002": "A_Brand",
        "SKU_003": "B_Brand",
        "SKU_004": "B_Brand",
    }

    today = datetime.today().date()
    dates = [today - timedelta(days=i) for i in range(59, -1, -1)]

    sales_rows = []
    for day in dates:
        for sku in skus:
            for channel in channels:
                base = 10 if channel == "D2C" else 7
                noise = np.random.poisson(2)
                qty = max(0, base + noise + np.random.randint(-3, 4))
                sales_rows.append(
                    {
                        "date": day,
                        "brand": sku_brand_map[sku],
                        "sku": sku,
                        "channel": channel,
                        "qty_sold": qty,
                    }
                )
    sales_df = pd.DataFrame(sales_rows)

    inventory_df = pd.DataFrame(
        {
            "date": [today] * len(skus),
            "sku": skus,
            "on_hand": [350, 180, 260, 120],
            "inbound_qty": [50, 40, 0, 30],
        }
    )

    policy_df = pd.DataFrame(
        {
            "sku": skus,
            "lead_time_days": [10, 14, 12, 9],
            "moq": [50, 50, 30, 30],
            "pack_size": [10, 10, 6, 6],
            "safety_stock": [80, 90, 70, 60],
        }
    )

    return sales_df, inventory_df, policy_df


def read_uploaded_excel(uploaded_file) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Read expected sheets from uploaded Excel file."""
    sales_df = pd.read_excel(uploaded_file, sheet_name="sales")
    inventory_df = pd.read_excel(uploaded_file, sheet_name="inventory")
    policy_df = pd.read_excel(uploaded_file, sheet_name="sku_policy")
    return sales_df, inventory_df, policy_df


def to_excel_bytes(
    sales_df: pd.DataFrame, inventory_df: pd.DataFrame, policy_df: pd.DataFrame
) -> bytes:
    """Create a template/sample xlsx for user download."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        sales_df.to_excel(writer, sheet_name="sales", index=False)
        inventory_df.to_excel(writer, sheet_name="inventory", index=False)
        policy_df.to_excel(writer, sheet_name="sku_policy", index=False)
    buffer.seek(0)
    return buffer.read()

