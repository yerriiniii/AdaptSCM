import uuid

from fastapi import APIRouter, Body, Depends, HTTPException, Query
from sqlalchemy import select
from sqlalchemy.orm import Session, selectinload

from domains.auth.dependencies import require_active_user
from domains.order.schemas import (
    PurchaseInboundCreateRequest,
    PurchaseInboundLineQuickPatchRequest,
    PurchaseInboundLineResponse,
    PurchaseOrderCreateRequest,
    PurchaseOrderImportRequest,
    PurchaseOrderImportResponse,
    PurchaseOrderListResponse,
    PurchaseOrderResponse,
    PurchaseOrderUpdateRequest,
)
from domains.order.models import PurchaseOrder as PurchaseOrderModel
from domains.order.services.purchase_order import (
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
from domains.item.schemas import SkuLookupForPurchaseResponse
from domains.item.services.item_mapping import find_item_by_any_country_sku
from shared.config import get_runtime_settings
from shared.db import get_db_session

router = APIRouter(
    prefix="/api/inventory/purchase-orders",
    tags=["order"],
    dependencies=[Depends(require_active_user)],
)


def _require_db_configured() -> None:
    if not get_runtime_settings().database_enabled:
        raise HTTPException(status_code=503, detail="DATABASE_URL is not configured.")


def _parse_uuid(value: str, label: str) -> uuid.UUID:
    try:
        return uuid.UUID(value)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"{label} format is invalid.") from exc


def _inbound_line_response(line, *, hide_expected_when_open: bool = True) -> PurchaseInboundLineResponse:
    return PurchaseInboundLineResponse(
        id=str(line.id),
        line_no=line.line_no,
        ref_code=line.ref_code,
        delivery_available_date=line.delivery_available_date.isoformat() if line.delivery_available_date else None,
        expected_inbound_date=(
            None
            if hide_expected_when_open and (line.inbound_status or "").strip().upper() == "O"
            else (line.expected_inbound_date.isoformat() if line.expected_inbound_date else None)
        ),
        actual_inbound_date=line.actual_inbound_date.isoformat() if line.actual_inbound_date else None,
        actual_inbound_note=line.actual_inbound_note,
        line_memo=line.line_memo,
        quantity=float(line.quantity),
        inbound_status=line.inbound_status,
        created_at=line.created_at.isoformat() if line.created_at else None,
    )


@router.get("/sku-hint", response_model=SkuLookupForPurchaseResponse)
def purchase_order_sku_hint(sku: str = Query(..., min_length=1), db: Session = Depends(get_db_session)) -> SkuLookupForPurchaseResponse:
    _require_db_configured()
    return SkuLookupForPurchaseResponse(**find_item_by_any_country_sku(db, sku))


@router.get("", response_model=PurchaseOrderListResponse)
def purchase_orders_list(db: Session = Depends(get_db_session)) -> PurchaseOrderListResponse:
    _require_db_configured()
    rows = list_purchase_orders(db)
    return PurchaseOrderListResponse(items=[PurchaseOrderResponse(**purchase_order_to_dict(po)) for po in rows])


@router.post("", response_model=PurchaseOrderResponse)
def purchase_orders_create(
    payload: PurchaseOrderCreateRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseOrderResponse:
    _require_db_configured()
    po = create_purchase_order(db, payload.model_dump())
    stmt = select(PurchaseOrderModel).where(PurchaseOrderModel.id == po.id).options(selectinload(PurchaseOrderModel.inbound_lines))
    po2 = db.scalars(stmt).unique().first()
    return PurchaseOrderResponse(**purchase_order_to_dict(po2 or po))


@router.post("/import", response_model=PurchaseOrderImportResponse)
def purchase_orders_import(
    payload: PurchaseOrderImportRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseOrderImportResponse:
    _require_db_configured()
    created = import_purchase_orders_from_sheet(db, [r.model_dump() for r in payload.rows])
    return PurchaseOrderImportResponse(created_count=created)


@router.delete("/{order_id}")
def purchase_orders_delete(order_id: str, db: Session = Depends(get_db_session)) -> dict:
    _require_db_configured()
    delete_purchase_order(db, _parse_uuid(order_id, "order_id"))
    return {"ok": True}


@router.patch("/{order_id}", response_model=PurchaseOrderResponse)
def purchase_orders_update(
    order_id: str,
    payload: PurchaseOrderUpdateRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseOrderResponse:
    _require_db_configured()
    po = update_purchase_order(db, _parse_uuid(order_id, "order_id"), payload.model_dump())
    stmt = (
        select(PurchaseOrderModel)
        .where(PurchaseOrderModel.id == po.id)
        .options(selectinload(PurchaseOrderModel.inbound_lines))
    )
    po2 = db.scalars(stmt).unique().first()
    return PurchaseOrderResponse(**purchase_order_to_dict(po2 or po))


@router.patch("/{order_id}/inbound-lines/{line_id}", response_model=PurchaseInboundLineResponse)
def purchase_orders_patch_inbound_line(
    order_id: str,
    line_id: str,
    payload: PurchaseInboundLineQuickPatchRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseInboundLineResponse:
    _require_db_configured()
    line = patch_inbound_line_quick(
        db,
        _parse_uuid(order_id, "order_id"),
        _parse_uuid(line_id, "line_id"),
        payload.model_dump(exclude_unset=True),
    )
    return _inbound_line_response(line)


@router.delete("/{order_id}/inbound-lines/{line_id}")
def purchase_orders_delete_inbound_line(order_id: str, line_id: str, db: Session = Depends(get_db_session)) -> dict:
    _require_db_configured()
    delete_inbound_line(db, _parse_uuid(order_id, "order_id"), _parse_uuid(line_id, "line_id"))
    return {"ok": True}


@router.post("/{order_id}/inbounds", response_model=PurchaseInboundLineResponse)
def purchase_orders_add_inbound(
    order_id: str,
    payload: PurchaseInboundCreateRequest = Body(...),
    db: Session = Depends(get_db_session),
) -> PurchaseInboundLineResponse:
    _require_db_configured()
    return _inbound_line_response(
        add_inbound_line(db, _parse_uuid(order_id, "order_id"), payload.model_dump()),
        hide_expected_when_open=False,
    )
