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
    ItemMasterExtraColumnCreateRequest,
    ItemMasterExtraColumnResponse,
    ItemMasterRowListResponse,
    ItemMasterRowPatchRequest,
    ItemMasterRowResponse,
    ItemMasterUploadResponse,
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
from domains.item.services.item_master import (
    MASTER_COLUMN_LABELS,
    MASTER_COLUMN_ORDER,
    add_item_master_extra_column,
    delete_item_master_extra_column,
    list_item_master_rows,
    merge_item_master_uploads,
    patch_item_master_row,
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


@router.get("/mappings/master-rows", response_model=ItemMasterRowListResponse)
def inventory_master_rows(
    query: str | None = Query(default=None),
    limit: int = Query(default=10000, ge=1, le=20000),
    db: Session = Depends(get_db_session),
) -> ItemMasterRowListResponse:
    items, extra_columns = list_item_master_rows(db=db, query=query, limit=limit)
    return ItemMasterRowListResponse(
        items=items,
        columns=[MASTER_COLUMN_LABELS[key] for key in MASTER_COLUMN_ORDER],
        extra_columns=extra_columns,
    )


@router.post("/mappings/master-rows/extra-columns", response_model=ItemMasterExtraColumnResponse)
def inventory_master_extra_column_add(
    payload: ItemMasterExtraColumnCreateRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> ItemMasterExtraColumnResponse:
    return ItemMasterExtraColumnResponse(**add_item_master_extra_column(db=db, label=payload.label))


@router.delete("/mappings/master-rows/extra-columns/{field_key}")
def inventory_master_extra_column_delete(
    field_key: str,
    db: Session = Depends(get_db_session),
) -> dict:
    return delete_item_master_extra_column(db=db, field_key=field_key)


@router.post("/mappings/master-rows/upload", response_model=ItemMasterUploadResponse)
def inventory_master_rows_upload(
    files: list[UploadFile] = File(...),
    db: Session = Depends(get_db_session),
) -> ItemMasterUploadResponse:
    result = merge_item_master_uploads(db=db, upload_files=files)
    return ItemMasterUploadResponse(**result)


@router.patch("/mappings/master-rows", response_model=ItemMasterRowResponse)
def inventory_master_row_patch(
    payload: ItemMasterRowPatchRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> ItemMasterRowResponse:
    return ItemMasterRowResponse(**patch_item_master_row(db=db, payload=payload.model_dump()))


__all__ = ["router"]
