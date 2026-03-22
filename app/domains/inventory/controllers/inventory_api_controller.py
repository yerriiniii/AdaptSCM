from fastapi import APIRouter, File, Form, UploadFile

from app.domains.inventory.dto.inventory_api_dto import InventoryAggregateResponse
from app.domains.inventory.services.inventory_aggregate_service import (
    aggregate_inventory_files,
    inspect_inventory_files,
)

router = APIRouter(prefix="/api/inventory", tags=["inventory"])


@router.post("/aggregate", response_model=InventoryAggregateResponse)
def aggregate_inventory(
    files: list[UploadFile] = File(...),
    file_dates: list[str] | None = Form(default=None),
    file_countries: list[str] | None = Form(default=None),
    keyword: str | None = Form(default=None),
    start_date: str | None = Form(default=None),
    end_date: str | None = Form(default=None),
    date_range: str | None = Form(default=None),
    level_filter: str | None = Form(default="1"),
) -> InventoryAggregateResponse:
    rows, dates = aggregate_inventory_files(
        files=files,
        keyword=keyword,
        start_date=start_date,
        end_date=end_date,
        date_range=date_range,
        level_filter=level_filter,
        file_dates=file_dates,
        file_countries=file_countries,
    )
    countries = sorted({str(row.get("country", "KR")) for row in rows})
    summary = {
        "item_count": len(rows),
        "date_count": len(dates),
    }
    return InventoryAggregateResponse(summary=summary, countries=countries, dates=dates, rows=rows)


@router.post("/file-metadata")
def inventory_file_metadata(
    files: list[UploadFile] = File(...),
    file_countries: list[str] | None = Form(default=None),
) -> dict:
    items = inspect_inventory_files(files=files, file_countries=file_countries)
    return {"files": items}

