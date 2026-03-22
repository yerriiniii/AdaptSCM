from pydantic import BaseModel


class InventoryAggregateResponse(BaseModel):
    summary: dict[str, int]
    countries: list[str]
    dates: list[str]
    rows: list[dict]

