from datetime import date, timedelta

import numpy as np
import pandas as pd


def build_sales_df(days: int = 120) -> pd.DataFrame:
    rng = np.random.default_rng(20260317)
    end_date = date.today()
    start_date = end_date - timedelta(days=days - 1)
    dates = pd.date_range(start=start_date, end=end_date, freq="D")

    sku_info = [
        ("A_Brand", "A_SKU_001"),
        ("A_Brand", "A_SKU_002"),
        ("A_Brand", "A_SKU_003"),
        ("A_Brand", "A_SKU_004"),
        ("B_Brand", "B_SKU_001"),
        ("B_Brand", "B_SKU_002"),
        ("B_Brand", "B_SKU_003"),
        ("B_Brand", "B_SKU_004"),
    ]
    channels = ["D2C", "OliveYoung", "NaverSmartStore"]

    rows = []
    for dt in dates:
        weekday_boost = 1.15 if dt.weekday() in (4, 5) else 1.0
        seasonal = 1.0 + 0.08 * np.sin((dt.dayofyear / 365) * 2 * np.pi)

        for brand, sku in sku_info:
            sku_base = 7 + int(sku[-1]) * 2
            for channel in channels:
                channel_factor = {"D2C": 1.2, "OliveYoung": 1.0, "NaverSmartStore": 0.85}[channel]
                demand_mean = sku_base * channel_factor * weekday_boost * seasonal
                qty = max(0, int(rng.poisson(demand_mean)))
                rows.append(
                    {
                        "date": dt.date(),
                        "brand": brand,
                        "sku": sku,
                        "channel": channel,
                        "qty_sold": qty,
                    }
                )
    return pd.DataFrame(rows)


def build_inventory_df() -> pd.DataFrame:
    today = date.today()
    skus = [
        "A_SKU_001",
        "A_SKU_002",
        "A_SKU_003",
        "A_SKU_004",
        "B_SKU_001",
        "B_SKU_002",
        "B_SKU_003",
        "B_SKU_004",
    ]
    on_hand = [420, 180, 95, 60, 310, 140, 85, 40]
    inbound_qty = [80, 0, 50, 30, 60, 40, 0, 20]
    return pd.DataFrame(
        {"date": [today] * len(skus), "sku": skus, "on_hand": on_hand, "inbound_qty": inbound_qty}
    )


def build_policy_df() -> pd.DataFrame:
    return pd.DataFrame(
        {
            "sku": [
                "A_SKU_001",
                "A_SKU_002",
                "A_SKU_003",
                "A_SKU_004",
                "B_SKU_001",
                "B_SKU_002",
                "B_SKU_003",
                "B_SKU_004",
            ],
            "lead_time_days": [12, 10, 15, 14, 11, 9, 16, 13],
            "moq": [60, 40, 30, 30, 50, 40, 30, 30],
            "pack_size": [10, 10, 6, 6, 10, 10, 6, 6],
            "safety_stock": [90, 75, 70, 65, 85, 70, 60, 55],
        }
    )


def write_standard_file(path: str) -> None:
    sales_df = build_sales_df()
    inventory_df = build_inventory_df()
    policy_df = build_policy_df()
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        sales_df.to_excel(writer, sheet_name="sales", index=False)
        inventory_df.to_excel(writer, sheet_name="inventory", index=False)
        policy_df.to_excel(writer, sheet_name="sku_policy", index=False)


def write_mapping_test_file(path: str) -> None:
    sales_df = build_sales_df().rename(
        columns={
            "date": "판매일",
            "brand": "브랜드명",
            "sku": "상품코드",
            "channel": "플랫폼",
            "qty_sold": "판매수량",
        }
    )
    inventory_df = build_inventory_df().rename(
        columns={"date": "기준일", "sku": "품목코드", "on_hand": "현재고", "inbound_qty": "입고예정수량"}
    )
    policy_df = build_policy_df().rename(
        columns={
            "sku": "제품코드",
            "lead_time_days": "리드타임",
            "moq": "최소주문수량",
            "pack_size": "박스단위",
            "safety_stock": "안전재고",
        }
    )
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        sales_df.to_excel(writer, sheet_name="판매데이터", index=False)
        inventory_df.to_excel(writer, sheet_name="재고현황", index=False)
        policy_df.to_excel(writer, sheet_name="상품정책", index=False)


if __name__ == "__main__":
    write_standard_file("sample_data/inventory_sample_standard.xlsx")
    write_mapping_test_file("sample_data/inventory_sample_mapping_test.xlsx")
    print("샘플 파일 생성 완료:")
    print("- sample_data/inventory_sample_standard.xlsx")
    print("- sample_data/inventory_sample_mapping_test.xlsx")
