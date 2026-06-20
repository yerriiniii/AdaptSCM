from fastapi import APIRouter, Body, Depends, File, Query, UploadFile
from sqlalchemy.orm import Session

from domains.auth.dependencies import require_active_user
from domains.item.schemas import (
    InventorySkuMappingItemPatchRequest,
    InventorySkuMappingItemResponse,
    InventorySkuMappingListResponse,
    InventorySkuMappingSegmentPatchRequest,
    InventorySkuMappingSummaryResponse,
    InventorySkuMappingUploadResponse,
    InventorySkuMappingUpsertRequest,
)
from domains.item.services.item_mapping import (
    clear_item_sku_mappings,
    get_item_sku_mapping_summary,
    list_item_sku_mappings,
    merge_item_sku_mappings,
    patch_item_sku_mapping_item,
    patch_item_sku_mapping_segment,
    upsert_item_sku_mapping,
)
from shared.db import get_db_session

router = APIRouter(
    prefix="/api/inventory",
    tags=["item"],
    dependencies=[Depends(require_active_user)],
)


@router.get("/mappings", response_model=InventorySkuMappingSummaryResponse)
def inventory_mappings(db: Session = Depends(get_db_session)) -> InventorySkuMappingSummaryResponse:
    return InventorySkuMappingSummaryResponse(**get_item_sku_mapping_summary(db))


@router.get("/mappings/items", response_model=InventorySkuMappingListResponse)
def inventory_mapping_items(
    query: str | None = Query(default=None),
    limit: int = Query(default=100, ge=1, le=5000),
    db: Session = Depends(get_db_session),
) -> InventorySkuMappingListResponse:
    return InventorySkuMappingListResponse(items=list_item_sku_mappings(db=db, query=query, limit=limit))


@router.post("/mappings/upload", response_model=InventorySkuMappingUploadResponse)
def inventory_mappings_upload(
    files: list[UploadFile] = File(...),
    db: Session = Depends(get_db_session),
) -> InventorySkuMappingUploadResponse:
    return InventorySkuMappingUploadResponse(**merge_item_sku_mappings(db=db, upload_files=files))


@router.post("/mappings/item", response_model=InventorySkuMappingItemResponse)
def inventory_mapping_upsert(
    payload: InventorySkuMappingUpsertRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> InventorySkuMappingItemResponse:
    return InventorySkuMappingItemResponse(**upsert_item_sku_mapping(db=db, payload=payload.model_dump()))


@router.patch("/mappings/item", response_model=InventorySkuMappingItemResponse)
def inventory_mapping_patch_item(
    payload: InventorySkuMappingItemPatchRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> InventorySkuMappingItemResponse:
    return InventorySkuMappingItemResponse(
        **patch_item_sku_mapping_item(db=db, payload=payload.model_dump(exclude_unset=True))
    )


@router.patch("/mappings/segment", response_model=InventorySkuMappingItemResponse)
def inventory_mapping_patch_segment(
    payload: InventorySkuMappingSegmentPatchRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> InventorySkuMappingItemResponse:
    return InventorySkuMappingItemResponse(
        **patch_item_sku_mapping_segment(
            db=db,
            group_id=payload.group_id,
            discontinued=payload.discontinued,
        )
    )


@router.delete("/mappings")
def inventory_mappings_clear(db: Session = Depends(get_db_session)) -> dict:
    return clear_item_sku_mappings(db=db)

__all__ = ["router"]
