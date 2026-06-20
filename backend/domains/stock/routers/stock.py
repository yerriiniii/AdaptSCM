from fastapi import APIRouter, Body, Depends, File, Form, Query, UploadFile
from sqlalchemy.orm import Session

from domains.auth.dependencies import require_active_user
from domains.stock.schemas import (
    InventoryAggregateResponse,
    InventoryCompleteUploadRequest,
    InventoryDirectUploadPreparedResponse,
    InventoryDirectUploadRequest,
    InventoryFilePatchBaseDateRequest,
)
from domains.stock.services.stock_aggregate import aggregate_stock_files, inspect_stock_files
from domains.stock.services.stock_persistence import (
    complete_stock_direct_uploads,
    delete_stock_file,
    delete_stock_files,
    get_stock_view,
    list_stock_files,
    patch_stock_file_base_date,
    persist_stock_uploads,
    prepare_stock_direct_uploads,
)
from shared.db import get_db_session

router = APIRouter(
    prefix="/api/inventory",
    tags=["stock"],
    dependencies=[Depends(require_active_user)],
)


@router.post("/aggregate", response_model=InventoryAggregateResponse)
def aggregate_inventory(
    files: list[UploadFile] = File(...),
    file_dates: list[str] | None = Form(default=None),
    file_countries: list[str] | None = Form(default=None),
    keyword: str | None = Form(default=None),
    start_date: str | None = Form(default=None),
    end_date: str | None = Form(default=None),
    date_range: str | None = Form(default=None),
    level_filter: str | None = Form(default="all"),
    file_client_ids: list[str] | None = Form(default=None),
    file_db_ids: list[str] | None = Form(default=None),
    db: Session = Depends(get_db_session),
) -> InventoryAggregateResponse:
    rows, dates = aggregate_stock_files(
        files=files,
        keyword=keyword,
        start_date=start_date,
        end_date=end_date,
        date_range=date_range,
        level_filter=level_filter,
        file_dates=file_dates,
        file_countries=file_countries,
        db=db,
    )
    persisted_files = persist_stock_uploads(
        db=db,
        files=files,
        file_dates=file_dates,
        file_countries=file_countries,
        file_client_ids=file_client_ids,
        file_db_ids=file_db_ids,
    )
    countries = sorted({str(row.get("country", "KR")) for row in rows})
    return InventoryAggregateResponse(
        summary={"item_count": len(rows), "date_count": len(dates)},
        countries=countries,
        dates=dates,
        rows=rows,
        files=persisted_files,
    )


@router.post("/file-metadata")
def inventory_file_metadata(
    files: list[UploadFile] = File(...),
    file_countries: list[str] | None = Form(default=None),
) -> dict:
    return {"files": inspect_stock_files(files=files, file_countries=file_countries)}


@router.post("/upload-url", response_model=InventoryDirectUploadPreparedResponse)
def inventory_upload_url(
    payload: InventoryDirectUploadRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> InventoryDirectUploadPreparedResponse:
    files = prepare_stock_direct_uploads([item.model_dump() for item in payload.files], db=db)
    return InventoryDirectUploadPreparedResponse(files=files)


@router.post("/complete-upload")
def inventory_complete_upload(
    payload: InventoryCompleteUploadRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> dict:
    return {"files": complete_stock_direct_uploads(db=db, files=[item.model_dump() for item in payload.files])}


@router.get("/files")
def inventory_files(db: Session = Depends(get_db_session)) -> dict:
    return {"files": list_stock_files(db)}


@router.patch("/files/{file_id}")
def inventory_patch_file(
    file_id: str,
    payload: InventoryFilePatchBaseDateRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> dict:
    return {"file": patch_stock_file_base_date(db=db, file_id=file_id, base_date_value=payload.base_date)}


@router.delete("/files/{file_id}")
def remove_inventory_file(file_id: str, db: Session = Depends(get_db_session)) -> dict:
    delete_stock_file(db=db, file_id=file_id)
    return {"ok": True}


@router.delete("/files")
def remove_inventory_files(
    country_code: str | None = Query(default=None),
    db: Session = Depends(get_db_session),
) -> dict:
    return delete_stock_files(db=db, country_code=country_code)


@router.get("/view", response_model=InventoryAggregateResponse)
def inventory_view(country_code: str = Query(...), db: Session = Depends(get_db_session)) -> InventoryAggregateResponse:
    return InventoryAggregateResponse(**get_stock_view(db=db, country_code=country_code))
