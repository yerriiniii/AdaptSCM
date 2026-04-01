from fastapi import APIRouter, Body, Depends, File, Form, Query, UploadFile
from sqlalchemy.orm import Session

from app.domains.inventory.dto.inventory_api_dto import (
    InventoryAggregateResponse,
    InventoryCompleteUploadRequest,
    InventoryDirectUploadPreparedResponse,
    InventoryDirectUploadRequest,
    InventorySkuMappingListResponse,
    InventorySkuMappingItemResponse,
    InventorySkuMappingSummaryResponse,
    InventorySkuMappingUpsertRequest,
    InventorySkuMappingUploadResponse,
)
from app.domains.inventory.services.inventory_mapping_service import (
    clear_product_sku_mappings,
    get_product_sku_mapping_summary,
    list_product_sku_mappings,
    replace_product_sku_mappings,
    upsert_product_sku_mapping,
)
from app.domains.inventory.services.inventory_aggregate_service import (
    aggregate_inventory_files,
    inspect_inventory_files,
)
from app.domains.inventory.services.inventory_persistence_service import (
    complete_inventory_direct_uploads,
    delete_inventory_file,
    delete_inventory_files,
    get_inventory_view,
    list_inventory_files,
    prepare_inventory_direct_uploads,
    persist_inventory_uploads,
)
from app.shared.db import get_db_session

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
    file_client_ids: list[str] | None = Form(default=None),
    file_db_ids: list[str] | None = Form(default=None),
    db: Session = Depends(get_db_session),
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
    persisted_files = persist_inventory_uploads(
        db=db,
        files=files,
        file_dates=file_dates,
        file_countries=file_countries,
        file_client_ids=file_client_ids,
        file_db_ids=file_db_ids,
    )
    countries = sorted({str(row.get("country", "KR")) for row in rows})
    summary = {
        "item_count": len(rows),
        "date_count": len(dates),
    }
    return InventoryAggregateResponse(
        summary=summary,
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
    items = inspect_inventory_files(files=files, file_countries=file_countries)
    return {"files": items}


@router.post("/upload-url", response_model=InventoryDirectUploadPreparedResponse)
def inventory_upload_url(
    payload: InventoryDirectUploadRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> InventoryDirectUploadPreparedResponse:
    return InventoryDirectUploadPreparedResponse(files=prepare_inventory_direct_uploads([item.model_dump() for item in payload.files], db=db))


@router.post("/complete-upload")
def inventory_complete_upload(
    payload: InventoryCompleteUploadRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> dict:
    return {"files": complete_inventory_direct_uploads(db=db, files=[item.model_dump() for item in payload.files])}


@router.get("/files")
def inventory_files(db: Session = Depends(get_db_session)) -> dict:
    return {"files": list_inventory_files(db)}


@router.get("/view", response_model=InventoryAggregateResponse)
def inventory_view(country_code: str = Query(...), db: Session = Depends(get_db_session)) -> InventoryAggregateResponse:
    payload = get_inventory_view(db=db, country_code=country_code)
    return InventoryAggregateResponse(**payload)


@router.get("/mappings", response_model=InventorySkuMappingSummaryResponse)
def inventory_mappings(db: Session = Depends(get_db_session)) -> InventorySkuMappingSummaryResponse:
    return InventorySkuMappingSummaryResponse(**get_product_sku_mapping_summary(db))


@router.get("/mappings/items", response_model=InventorySkuMappingListResponse)
def inventory_mapping_items(
    query: str | None = Query(default=None),
    limit: int = Query(default=100, ge=1, le=500),
    db: Session = Depends(get_db_session),
) -> InventorySkuMappingListResponse:
    return InventorySkuMappingListResponse(items=list_product_sku_mappings(db=db, query=query, limit=limit))


@router.post("/mappings/upload", response_model=InventorySkuMappingUploadResponse)
def inventory_mappings_upload(
    file: UploadFile = File(...),
    db: Session = Depends(get_db_session),
) -> InventorySkuMappingUploadResponse:
    return InventorySkuMappingUploadResponse(**replace_product_sku_mappings(db=db, upload_file=file))


@router.post("/mappings/item", response_model=InventorySkuMappingItemResponse)
def inventory_mapping_upsert(
    payload: InventorySkuMappingUpsertRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> InventorySkuMappingItemResponse:
    return InventorySkuMappingItemResponse(**upsert_product_sku_mapping(db=db, payload=payload.model_dump()))


@router.delete("/mappings")
def inventory_mappings_clear(db: Session = Depends(get_db_session)) -> dict:
    return clear_product_sku_mappings(db=db)


@router.delete("/files/{file_id}")
def remove_inventory_file(file_id: str, db: Session = Depends(get_db_session)) -> dict:
    delete_inventory_file(db=db, file_id=file_id)
    return {"ok": True}


@router.delete("/files")
def remove_inventory_files(
    country_code: str | None = Query(default=None),
    db: Session = Depends(get_db_session),
) -> dict:
    return delete_inventory_files(db=db, country_code=country_code)

