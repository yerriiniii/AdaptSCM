"""발주 기록 CRUD."""

from __future__ import annotations

import uuid
from datetime import date
from decimal import Decimal

from fastapi import HTTPException
from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from app.domains.inventory.models.inventory_models import PurchaseInboundLine, PurchaseOrder


def _parse_date(value: object, field: str) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s:
        return None
    try:
        return date.fromisoformat(s[:10])
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=f"{field} 날짜 형식이 올바르지 않습니다. (YYYY-MM-DD)") from exc


def _parse_decimal(value: object, field: str) -> Decimal:
    try:
        d = Decimal(str(value))
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"{field} 숫자 형식이 올바르지 않습니다.") from exc
    return d


def list_purchase_orders(db: Session) -> list[PurchaseOrder]:
    stmt = (
        select(PurchaseOrder)
        .options(selectinload(PurchaseOrder.inbound_lines))
        .order_by(PurchaseOrder.created_at.desc())
    )
    return list(db.scalars(stmt).unique().all())


def get_purchase_order(db: Session, order_id: uuid.UUID) -> PurchaseOrder | None:
    stmt = (
        select(PurchaseOrder)
        .where(PurchaseOrder.id == order_id)
        .options(selectinload(PurchaseOrder.inbound_lines))
    )
    return db.scalars(stmt).unique().first()


def create_purchase_order(db: Session, payload: dict) -> PurchaseOrder:
    erp = str(payload.get("erp_po_number") or "").strip()
    if not erp:
        raise HTTPException(status_code=400, detail="ERP PO 번호는 필수입니다.")
    existing = db.scalar(select(PurchaseOrder.id).where(PurchaseOrder.erp_po_number == erp))
    if existing:
        raise HTTPException(status_code=409, detail="이미 저장된 발주 기록입니다.")

    sku = str(payload.get("sku") or "").strip()
    if not sku:
        raise HTTPException(status_code=400, detail="상품코드(SKU)는 필수입니다.")

    order_date = _parse_date(payload.get("order_date"), "발주일자")
    if order_date is None:
        raise HTTPException(status_code=400, detail="발주일자는 필수입니다.")

    total_q = _parse_decimal(payload.get("total_quantity"), "총 발주수량")

    po = PurchaseOrder(
        order_date=order_date,
        erp_po_number=erp,
        product_type=str(payload.get("product_type") or "").strip() or "본품",
        sku=sku,
        brand=str(payload.get("brand") or "").strip(),
        product_name=str(payload.get("product_name") or "").strip(),
        manufacturer=str(payload.get("manufacturer") or "").strip(),
        total_quantity=float(total_q),
        delivery_available_date=_parse_date(payload.get("delivery_available_date"), "납품가능일"),
        expected_inbound_date=_parse_date(payload.get("expected_inbound_date"), "입고예정일"),
    )
    db.add(po)
    db.commit()
    db.refresh(po)
    return po


def add_inbound_line(db: Session, order_id: uuid.UUID, payload: dict) -> PurchaseInboundLine:
    po = get_purchase_order(db, order_id)
    if po is None:
        raise HTTPException(status_code=404, detail="발주를 찾을 수 없습니다.")

    max_no = db.scalar(
        select(func.max(PurchaseInboundLine.line_no)).where(PurchaseInboundLine.purchase_order_id == po.id)
    )
    next_no = int(max_no or 0) + 1
    ref = f"{po.erp_po_number}_{next_no}"

    qty = _parse_decimal(payload.get("quantity"), "입고수량")
    status = str(payload.get("inbound_status") or "O").strip().upper()
    if status not in ("O", "X"):
        raise HTTPException(status_code=400, detail="입고여부는 O 또는 X 여야 합니다.")

    line = PurchaseInboundLine(
        purchase_order_id=po.id,
        line_no=next_no,
        ref_code=ref,
        actual_inbound_date=_parse_date(payload.get("actual_inbound_date"), "실제입고일"),
        quantity=float(qty),
        inbound_status=status,
    )
    db.add(line)
    db.commit()
    db.refresh(line)
    return line


def update_purchase_order(db: Session, order_id: uuid.UUID, payload: dict) -> PurchaseOrder:
    po = get_purchase_order(db, order_id)
    if po is None:
        raise HTTPException(status_code=404, detail="발주를 찾을 수 없습니다.")

    new_erp = str(payload.get("erp_po_number") or "").strip()
    if not new_erp:
        raise HTTPException(status_code=400, detail="ERP PO 번호는 필수입니다.")

    conflict = db.scalar(
        select(PurchaseOrder.id).where(
            PurchaseOrder.erp_po_number == new_erp,
            PurchaseOrder.id != order_id,
        )
    )
    if conflict:
        raise HTTPException(status_code=409, detail="이미 저장된 발주 기록입니다.")

    sku = str(payload.get("sku") or "").strip()
    if not sku:
        raise HTTPException(status_code=400, detail="상품코드(SKU)는 필수입니다.")

    order_date = _parse_date(payload.get("order_date"), "발주일자")
    if order_date is None:
        raise HTTPException(status_code=400, detail="발주일자는 필수입니다.")

    total_q = _parse_decimal(payload.get("total_quantity"), "총 발주수량")

    old_erp = po.erp_po_number
    po.order_date = order_date
    po.erp_po_number = new_erp
    po.product_type = str(payload.get("product_type") or "").strip() or "본품"
    po.sku = sku
    po.brand = str(payload.get("brand") or "").strip()
    po.product_name = str(payload.get("product_name") or "").strip()
    po.manufacturer = str(payload.get("manufacturer") or "").strip()
    po.total_quantity = float(total_q)
    po.delivery_available_date = _parse_date(payload.get("delivery_available_date"), "납품가능일")
    po.expected_inbound_date = _parse_date(payload.get("expected_inbound_date"), "입고예정일")

    if old_erp != new_erp:
        for line in po.inbound_lines or []:
            line.ref_code = f"{new_erp}_{line.line_no}"

    line_updates = payload.get("inbound_lines")
    if line_updates is not None:
        by_id = {str(l.id): l for l in (po.inbound_lines or [])}
        for item in line_updates:
            lid = str(item.get("id") or "").strip()
            if not lid:
                raise HTTPException(status_code=400, detail="입고 차수 id가 없습니다.")
            line = by_id.get(lid)
            if line is None:
                raise HTTPException(status_code=400, detail=f"입고 차수를 찾을 수 없습니다: {lid}")
            qty = _parse_decimal(item.get("quantity"), "입고수량")
            status = str(item.get("inbound_status") or "O").strip().upper()
            if status not in ("O", "X"):
                raise HTTPException(status_code=400, detail="입고여부는 O 또는 X 여야 합니다.")
            line.actual_inbound_date = _parse_date(item.get("actual_inbound_date"), "실제입고일")
            line.quantity = float(qty)
            line.inbound_status = status

    db.commit()
    db.refresh(po)
    return po


def purchase_order_to_dict(po: PurchaseOrder) -> dict:
    lines = sorted(po.inbound_lines or [], key=lambda x: x.line_no)
    return {
        "id": str(po.id),
        "order_date": po.order_date.isoformat() if po.order_date else None,
        "erp_po_number": po.erp_po_number,
        "product_type": po.product_type,
        "sku": po.sku,
        "brand": po.brand,
        "product_name": po.product_name,
        "manufacturer": po.manufacturer,
        "total_quantity": float(po.total_quantity),
        "delivery_available_date": po.delivery_available_date.isoformat() if po.delivery_available_date else None,
        "expected_inbound_date": po.expected_inbound_date.isoformat() if po.expected_inbound_date else None,
        "created_at": po.created_at.isoformat() if po.created_at else None,
        "inbound_lines": [
            {
                "id": str(line.id),
                "line_no": line.line_no,
                "ref_code": line.ref_code,
                "actual_inbound_date": line.actual_inbound_date.isoformat() if line.actual_inbound_date else None,
                "quantity": float(line.quantity),
                "inbound_status": line.inbound_status,
                "created_at": line.created_at.isoformat() if line.created_at else None,
            }
            for line in lines
        ],
    }
