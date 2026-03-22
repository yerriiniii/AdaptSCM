import numpy as np
import pandas as pd

from app.domains.reorder.dto.reorder_dto import ReorderSettings


def compute_weighted_daily_sales(
    sales_df: pd.DataFrame, w7: float, w14: float, w28: float
) -> pd.DataFrame:
    """Compute weighted daily demand per sku using recent windows."""
    latest_date = sales_df["date"].max()

    def daily_avg(window_days: int) -> pd.Series:
        start = latest_date - pd.Timedelta(days=window_days - 1)
        mask = (sales_df["date"] >= start) & (sales_df["date"] <= latest_date)
        grouped = sales_df.loc[mask].groupby("sku")["qty_sold"].sum()
        return grouped / window_days

    avg7 = daily_avg(7)
    avg14 = daily_avg(14)
    avg28 = daily_avg(28)

    demand = pd.concat([avg7, avg14, avg28], axis=1)
    demand.columns = ["avg7", "avg14", "avg28"]
    demand = demand.fillna(0.0)
    demand["weighted_daily_demand"] = (
        demand["avg7"] * w7 + demand["avg14"] * w14 + demand["avg28"] * w28
    )
    return demand.reset_index()[["sku", "weighted_daily_demand", "avg7", "avg14", "avg28"]]


def apply_order_constraints(order_qty: float, moq: float, pack_size: float) -> int:
    """Apply MOQ and pack size rounding rules to recommendation."""
    if order_qty <= 0:
        return 0
    constrained = max(order_qty, moq)
    packs = np.ceil(constrained / pack_size)
    return int(packs * pack_size)


def calculate_reorder_plan(
    sales_df: pd.DataFrame,
    inventory_df: pd.DataFrame,
    policy_df: pd.DataFrame,
    settings: ReorderSettings,
) -> pd.DataFrame:
    demand_df = compute_weighted_daily_sales(
        sales_df, w7=settings.w7, w14=settings.w14, w28=settings.w28
    )

    inventory_latest = (
        inventory_df.sort_values("date")
        .groupby("sku", as_index=False)
        .tail(1)[["sku", "on_hand", "inbound_qty"]]
    )

    merged = (
        policy_df.merge(inventory_latest, on="sku", how="left")
        .merge(demand_df, on="sku", how="left")
        .fillna(
            {
                "on_hand": 0,
                "inbound_qty": 0,
                "weighted_daily_demand": 0,
                "avg7": 0,
                "avg14": 0,
                "avg28": 0,
            }
        )
    )

    merged["lead_time_demand"] = merged["weighted_daily_demand"] * merged["lead_time_days"]
    merged["reorder_point"] = merged["lead_time_demand"] + merged["safety_stock"]
    merged["available_stock"] = merged["on_hand"] + merged["inbound_qty"]
    merged["target_stock"] = (
        merged["weighted_daily_demand"] * settings.target_cover_days
        + merged["safety_stock"]
    )
    merged["raw_order_qty"] = merged["target_stock"] - merged["available_stock"]

    merged["recommended_order_qty"] = merged.apply(
        lambda row: apply_order_constraints(row["raw_order_qty"], row["moq"], row["pack_size"]),
        axis=1,
    )
    merged["days_of_supply"] = np.where(
        merged["weighted_daily_demand"] > 0,
        merged["on_hand"] / merged["weighted_daily_demand"],
        np.inf,
    )
    merged["need_reorder"] = merged["available_stock"] <= merged["reorder_point"]

    output_cols = [
        "sku",
        "weighted_daily_demand",
        "avg7",
        "avg14",
        "avg28",
        "on_hand",
        "inbound_qty",
        "days_of_supply",
        "lead_time_days",
        "safety_stock",
        "reorder_point",
        "target_stock",
        "raw_order_qty",
        "recommended_order_qty",
        "need_reorder",
    ]
    return merged[output_cols].sort_values(
        ["need_reorder", "recommended_order_qty"], ascending=[False, False]
    )

