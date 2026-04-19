import uuid

from fastapi import APIRouter, Body, Depends, File, Form, HTTPException, Query, UploadFile
from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

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
    PurchaseInboundCreateRequest,
    PurchaseInboundLineQuickPatchRequest,
    PurchaseInboundLineResponse,
    PurchaseOrderCreateRequest,
    PurchaseOrderImportRequest,
    PurchaseOrderImportResponse,
    PurchaseOrderListResponse,
    PurchaseOrderResponse,
    PurchaseOrderUpdateRequest,
    SkuLookupForPurchaseResponse,
)
from app.domains.inventory.services.inventory_mapping_service import (
    clear_product_sku_mappings,
    get_product_sku_mapping_summary,
    list_product_sku_mappings,
    lookup_product_by_any_country_sku,
    merge_product_sku_mappings,
    upsert_product_sku_mapping,
)
from app.domains.inventory.services.inventory_aggregate_service import (
    aggregate_inventory_files,
    inspect_inventory_files,
)
from app.domains.inventory.services.shipment_aggregate_service import (
    build_shipment_wide_dashboard,
    collect_shipment_long_records_from_raw,
)
from app.domains.inventory.services.shipment_persistence_service import get_shipment_view, persist_shipment_uploads
from app.domains.inventory.services.inventory_persistence_service import (
    _read_upload_bytes,
    complete_inventory_direct_uploads,
    delete_inventory_file,
    delete_inventory_files,
    get_inventory_view,
    list_inventory_files,
    prepare_inventory_direct_uploads,
    persist_inventory_uploads,
)
from app.domains.inventory.models.inventory_models import PurchaseOrder as PurchaseOrderModel
from app.domains.inventory.services.purchase_order_service import (
    add_inbound_line,
    create_purchase_order,
    delete_inbound_line,
    delete_purchase_order,
    import_purchase_orders_from_sheet,
    list_purchase_orders,
    patch_inbound_line_quick,
    purchase_order_to_dict,
    update_purchase_order,
)
from app.shared.config import get_runtime_settings
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
    level_filter: str | None = Form(default="all"),
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
        db=db,
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


@router.post("/shipment/aggregate", response_model=InventoryAggregateResponse)
def shipment_aggregate(
    files: list[UploadFile] = File(...),
    db: Session = Depends(get_db_session),
) -> InventoryAggregateResponse:
    file_payloads: list[tuple[UploadFile, bytes, list[dict]]] = []
    all_long: list[dict] = []
    for uf in files:
        raw = _read_upload_bytes(uf)
        recs = collect_shipment_long_records_from_raw(db, raw, uf.filename or "")
        all_long.extend(recs)
        file_payloads.append((uf, raw, recs))
    payload = build_shipment_wide_dashboard(all_long)
    persisted_files: list[dict] = []
    settings = get_runtime_settings()
    if settings.database_enabled:
        persisted_files = persist_shipment_uploads(db, file_payloads)
    return InventoryAggregateResponse(
        summary=payload["summary"],
        countries=payload["countries"],
        dates=payload["dates"],
        rows=payload["rows"],
        channels=payload["channels"],
        files=persisted_files,
    )


@router.get("/shipment/view", response_model=InventoryAggregateResponse)
def shipment_view(db: Session = Depends(get_db_session)) -> InventoryAggregateResponse:
    if not get_runtime_settings().database_enabled:
        raise HTTPException(status_code=503, detail="DATABASE_URL이 설정되지 않았습니다.")
    return InventoryAggregateResponse(**get_shipment_view(db))


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
    files: list[UploadFile] = File(...),
    db: Session = Depends(get_db_session),
) -> InventorySkuMappingUploadResponse:
    return InventorySkuMappingUploadResponse(**merge_product_sku_mappings(db=db, upload_files=files))


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


def _require_db_configured() -> None:
    if not get_runtime_settings().database_enabled:
        raise HTTPException(status_code=503, detail="DATABASE_URL이 설정되지 않았습니다.")


@router.get("/purchase-orders/sku-hint", response_model=SkuLookupForPurchaseResponse)
def purchase_order_sku_hint(sku: str = Query(..., min_length=1), db: Session = Depends(get_db_session)) -> SkuLookupForPurchaseResponse:
    _require_db_configured()
    data = lookup_product_by_any_country_sku(db, sku)
    return SkuLookupForPurchaseResponse(**data)


@router.get("/purchase-orders", response_model=PurchaseOrderListResponse)
def purchase_orders_list(db: Session = Depends(get_db_session)) -> PurchaseOrderListResponse:
    _require_db_configured()
    rows = list_purchase_orders(db)
    return PurchaseOrderListResponse(items=[PurchaseOrderResponse(**purchase_order_to_dict(po)) for po in rows])


@router.post("/purchase-orders", response_model=PurchaseOrderResponse)
def purchase_orders_create(
    payload: PurchaseOrderCreateRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseOrderResponse:
    _require_db_configured()
    po = create_purchase_order(db, payload.model_dump())
    stmt = select(PurchaseOrderModel).where(PurchaseOrderModel.id == po.id).options(selectinload(PurchaseOrderModel.inbound_lines))
    po2 = db.scalars(stmt).unique().first()
    return PurchaseOrderResponse(**purchase_order_to_dict(po2 or po))


@router.post("/purchase-orders/import", response_model=PurchaseOrderImportResponse)
def purchase_orders_import(
    payload: PurchaseOrderImportRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseOrderImportResponse:
    _require_db_configured()
    created = import_purchase_orders_from_sheet(db, [r.model_dump() for r in payload.rows])
    return PurchaseOrderImportResponse(created_count=created)


@router.delete("/purchase-orders/{order_id}")
def purchase_orders_delete(order_id: str, db: Session = Depends(get_db_session)) -> dict:
    _require_db_configured()
    try:
        oid = uuid.UUID(order_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="order_id 형식이 올바르지 않습니다.") from exc
    delete_purchase_order(db, oid)
    return {"ok": True}


@router.patch("/purchase-orders/{order_id}", response_model=PurchaseOrderResponse)
def purchase_orders_update(
    order_id: str,
    payload: PurchaseOrderUpdateRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseOrderResponse:
    _require_db_configured()
    try:
        oid = uuid.UUID(order_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="order_id 형식이 올바르지 않습니다.") from exc
    po = update_purchase_order(db, oid, payload.model_dump())
    stmt = (
        select(PurchaseOrderModel)
        .where(PurchaseOrderModel.id == po.id)
        .options(selectinload(PurchaseOrderModel.inbound_lines))
    )
    po2 = db.scalars(stmt).unique().first()
    return PurchaseOrderResponse(**purchase_order_to_dict(po2 or po))


@router.patch(
    "/purchase-orders/{order_id}/inbound-lines/{line_id}",
    response_model=PurchaseInboundLineResponse,
)
def purchase_orders_patch_inbound_line(
    order_id: str,
    line_id: str,
    payload: PurchaseInboundLineQuickPatchRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseInboundLineResponse:
    _require_db_configured()
    try:
        oid = uuid.UUID(order_id)
        lid = uuid.UUID(line_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="id 형식이 올바르지 않습니다.") from exc
    line = patch_inbound_line_quick(db, oid, lid, payload.model_dump(exclude_unset=True))
    return PurchaseInboundLineResponse(
        id=str(line.id),
        line_no=line.line_no,
        ref_code=line.ref_code,
        delivery_available_date=(
            line.delivery_available_date.isoformat() if line.delivery_available_date else None
        ),
        expected_inbound_date=(
            None
            if (line.inbound_status or "").strip().upper() == "O"
            else (line.expected_inbound_date.isoformat() if line.expected_inbound_date else None)
        ),
        actual_inbound_date=line.actual_inbound_date.isoformat() if line.actual_inbound_date else None,
        actual_inbound_note=line.actual_inbound_note,
        line_memo=line.line_memo,
        quantity=float(line.quantity),
        inbound_status=line.inbound_status,
        created_at=line.created_at.isoformat() if line.created_at else None,
    )


@router.delete("/purchase-orders/{order_id}/inbound-lines/{line_id}")
def purchase_orders_delete_inbound_line(order_id: str, line_id: str, db: Session = Depends(get_db_session)) -> dict:
    _require_db_configured()
    try:
        oid = uuid.UUID(order_id)
        lid = uuid.UUID(line_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="id 형식이 올바르지 않습니다.") from exc
    delete_inbound_line(db, oid, lid)
    return {"ok": True}


@router.post("/purchase-orders/{order_id}/inbounds", response_model=PurchaseInboundLineResponse)
def purchase_orders_add_inbound(
    order_id: str,
    payload: PurchaseInboundCreateRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseInboundLineResponse:
    _require_db_configured()
    try:
        oid = uuid.UUID(order_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail="order_id 형식이 올바르지 않습니다.") from exc
    line = add_inbound_line(db, oid, payload.model_dump())
    return PurchaseInboundLineResponse(
        id=str(line.id),
        line_no=line.line_no,
        ref_code=line.ref_code,
        delivery_available_date=(
            line.delivery_available_date.isoformat() if line.delivery_available_date else None
        ),
        expected_inbound_date=(
            line.expected_inbound_date.isoformat() if line.expected_inbound_date else None
        ),
        actual_inbound_date=line.actual_inbound_date.isoformat() if line.actual_inbound_date else None,
        actual_inbound_note=line.actual_inbound_note,
        line_memo=line.line_memo,
        quantity=float(line.quantity),
        inbound_status=line.inbound_status,
        created_at=line.created_at.isoformat() if line.created_at else None,
    )

