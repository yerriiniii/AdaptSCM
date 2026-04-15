"""발주 기록 CRUD."""

from __future__ import annotations

import uuid
from datetime import date
from decimal import Decimal

from fastapi import HTTPException
from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from app.domains.inventory.models.inventory_models import PurchaseInboundLine, PurchaseOrder

ORDER_DATE_PLANNED_CANONICAL = "발주 예정"
ORDER_DATE_NOTE_MAX = 128
NO_ERP_PO_PREFIX = "__NO_ERP__"


def _normalize_inbound_ref_codes(po: PurchaseOrder) -> None:
    """입고 차수 1개면 ERP PO만, 2개 이상이면 _1/_2(내부용 PO) 또는 ERP_1/ERP_2."""
    lines = sorted(po.inbound_lines or [], key=lambda x: x.line_no)
    erp = str(po.erp_po_number or "").strip()
    n = len(lines)
    if n == 0:
        return
    if n == 1:
        lines[0].ref_code = (erp[:160] if erp else "-")[:160]
        return
    no_internal = erp.startswith(NO_ERP_PO_PREFIX)
    for ln in lines:
        k = int(ln.line_no or 0)
        seg = f"_{k}" if no_internal else f"{erp}_{k}"
        ln.ref_code = seg[:160]


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


_INBOUND_NOTE_MAX = 128
_LINE_MEMO_MAX = 4000


def _parse_order_date_and_note_from_payload(payload: dict) -> tuple[date | None, str | None]:
    """발주일: 날짜가 있으면 그만 쓰고 비고는 비움. 날짜·비고 모두 비면 (None, None). 비고만 있으면 「발주 예정」만 허용."""
    note_raw = str(payload.get("order_date_note") or "").strip() or None
    if note_raw and len(note_raw) > ORDER_DATE_NOTE_MAX:
        raise HTTPException(
            status_code=400,
            detail=f"발주일 비고는 {ORDER_DATE_NOTE_MAX}자 이내여야 합니다.",
        )
    s_date = str(payload.get("order_date") or "").strip()
    if s_date:
        d = _parse_date(payload.get("order_date"), "발주일자")
        return d, None
    if note_raw:
        n = " ".join(note_raw.split())
        compact = note_raw.replace(" ", "").replace("\u3000", "")
        if n == ORDER_DATE_PLANNED_CANONICAL or compact == "발주예정":
            return None, ORDER_DATE_PLANNED_CANONICAL
        raise HTTPException(
            status_code=400,
            detail=f"발주일이 비어 있을 때 비고에는 「{ORDER_DATE_PLANNED_CANONICAL}」만 입력할 수 있습니다.",
        )
    return None, None


def _parse_inbound_date_and_note(payload: dict) -> tuple[date | None, str | None]:
    """날짜가 있으면 날짜만 사용. 없으면 비고(예: 예외 입고·무상 입고)만 저장할 수 있음."""
    note = str(payload.get("actual_inbound_note") or "").strip() or None
    if note and len(note) > _INBOUND_NOTE_MAX:
        raise HTTPException(
            status_code=400,
            detail=f"실제입고 비고는 {_INBOUND_NOTE_MAX}자 이내여야 합니다.",
        )
    ad = _parse_date(payload.get("actual_inbound_date"), "실제입고일")
    if ad is not None:
        return ad, None
    return None, note


def list_purchase_orders(db: Session) -> list[PurchaseOrder]:
    stmt = (
        select(PurchaseOrder)
        .options(selectinload(PurchaseOrder.inbound_lines))
        .order_by(PurchaseOrder.created_at.desc())
    )
    return list(db.scalars(stmt).unique().all())


def find_existing_purchase_order_id_by_business_key(
    db: Session,
    erp_po_number: str,
    sku: str,
    order_date: date | None,
    *,
    exclude_po_id: uuid.UUID | None = None,
) -> uuid.UUID | None:
    """동일 발주건: ERP PO + 상품코드(SKU) + 발주일(order_date, 미정이면 둘 다 NULL)."""
    stmt = select(PurchaseOrder.id).where(
        PurchaseOrder.erp_po_number == erp_po_number,
        PurchaseOrder.sku == sku,
    )
    if order_date is None:
        stmt = stmt.where(PurchaseOrder.order_date.is_(None))
    else:
        stmt = stmt.where(PurchaseOrder.order_date == order_date)
    if exclude_po_id is not None:
        stmt = stmt.where(PurchaseOrder.id != exclude_po_id)
    return db.scalar(stmt)


def get_purchase_order(db: Session, order_id: uuid.UUID) -> PurchaseOrder | None:
    stmt = (
        select(PurchaseOrder)
        .where(PurchaseOrder.id == order_id)
        .options(selectinload(PurchaseOrder.inbound_lines))
    )
    return db.scalars(stmt).unique().first()


def delete_purchase_order(db: Session, order_id: uuid.UUID) -> None:
    po = get_purchase_order(db, order_id)
    if po is None:
        raise HTTPException(status_code=404, detail="발주를 찾을 수 없습니다.")
    db.delete(po)
    db.commit()


def create_purchase_order(db: Session, payload: dict) -> PurchaseOrder:
    erp = str(payload.get("erp_po_number") or "").strip()
    if not erp:
        raise HTTPException(status_code=400, detail="ERP PO 번호는 필수입니다.")

    sku = str(payload.get("sku") or "").strip()
    if not sku:
        raise HTTPException(status_code=400, detail="상품코드(SKU)는 필수입니다.")

    od, onote = _parse_order_date_and_note_from_payload(payload)

    existing = find_existing_purchase_order_id_by_business_key(db, erp, sku, od)
    if existing:
        raise HTTPException(
            status_code=409,
            detail="동일 ERP PO·상품코드·발주일 조합의 발주가 이미 있습니다.",
        )

    total_q = _parse_decimal(payload.get("total_quantity"), "총 발주수량")

    po = PurchaseOrder(
        order_date=od,
        order_date_note=onote,
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

    qty = _parse_decimal(payload.get("quantity"), "입고수량")
    status = str(payload.get("inbound_status") or "O").strip().upper()
    if status not in ("O", "X"):
        raise HTTPException(status_code=400, detail="입고여부는 O 또는 X 여야 합니다.")

    ad, note = _parse_inbound_date_and_note(payload)
    if status == "O" and ad is None and not note:
        raise HTTPException(
            status_code=400,
            detail="입고 완료(O)이면 실제입고일 또는 실제입고 비고(예: 예외 입고·무상 입고)가 필요합니다.",
        )

    line_delivery = _parse_date(payload.get("delivery_available_date"), "납품가능일")
    line_expected = (
        None
        if status == "O"
        else _parse_date(payload.get("expected_inbound_date"), "입고예정일")
    )

    line = PurchaseInboundLine(
        purchase_order_id=po.id,
        line_no=next_no,
        ref_code="_",
        delivery_available_date=line_delivery,
        expected_inbound_date=line_expected,
        actual_inbound_date=ad,
        actual_inbound_note=note,
        quantity=float(qty),
        inbound_status=status,
    )
    db.add(line)
    db.flush()
    po2 = get_purchase_order(db, order_id)
    if po2 is not None:
        _normalize_inbound_ref_codes(po2)
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

    sku = str(payload.get("sku") or "").strip()
    if not sku:
        raise HTTPException(status_code=400, detail="상품코드(SKU)는 필수입니다.")

    od, onote = _parse_order_date_and_note_from_payload(payload)

    conflict = find_existing_purchase_order_id_by_business_key(
        db, new_erp, sku, od, exclude_po_id=order_id
    )
    if conflict:
        raise HTTPException(
            status_code=409,
            detail="동일 ERP PO·상품코드·발주일 조합의 발주가 이미 있습니다.",
        )

    total_q = _parse_decimal(payload.get("total_quantity"), "총 발주수량")

    old_erp = po.erp_po_number
    po.order_date = od
    po.order_date_note = onote
    po.erp_po_number = new_erp
    po.product_type = str(payload.get("product_type") or "").strip() or "본품"
    po.sku = sku
    po.brand = str(payload.get("brand") or "").strip()
    po.product_name = str(payload.get("product_name") or "").strip()
    po.manufacturer = str(payload.get("manufacturer") or "").strip()
    po.total_quantity = float(total_q)
    po.delivery_available_date = _parse_date(payload.get("delivery_available_date"), "납품가능일")
    po.expected_inbound_date = _parse_date(payload.get("expected_inbound_date"), "입고예정일")

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
            ad, note = _parse_inbound_date_and_note(item)
            if status == "O" and ad is None and not note:
                raise HTTPException(
                    status_code=400,
                    detail="입고 완료(O)이면 실제입고일 또는 실제입고 비고(예: 예외 입고·무상 입고)가 필요합니다.",
                )
            line.actual_inbound_date = ad
            line.actual_inbound_note = note
            line.quantity = float(qty)
            line.inbound_status = status
            if "line_memo" in item:
                m = str(item.get("line_memo") or "").strip() or None
                if m and len(m) > _LINE_MEMO_MAX:
                    raise HTTPException(
                        status_code=400,
                        detail=f"비고 메모는 {_LINE_MEMO_MAX}자 이내여야 합니다.",
                    )
                line.line_memo = m
            if "delivery_available_date" in item:
                line.delivery_available_date = _parse_date(
                    item.get("delivery_available_date"), "납품가능일"
                )
            if status == "O":
                line.expected_inbound_date = None
            elif "expected_inbound_date" in item:
                line.expected_inbound_date = _parse_date(
                    item.get("expected_inbound_date"), "입고예정일"
                )

    _normalize_inbound_ref_codes(po)

    db.commit()
    db.refresh(po)
    return po


def patch_inbound_line_quick(db: Session, order_id: uuid.UUID, line_id: uuid.UUID, payload: dict) -> PurchaseInboundLine:
    """실제입고일·입고수량·입고여부·비고 메모만 부분 수정."""
    po = get_purchase_order(db, order_id)
    if po is None:
        raise HTTPException(status_code=404, detail="발주를 찾을 수 없습니다.")
    line = next((ln for ln in (po.inbound_lines or []) if ln.id == line_id), None)
    if line is None:
        raise HTTPException(status_code=404, detail="입고 차수를 찾을 수 없습니다.")

    touched = False
    if "actual_inbound_date" in payload:
        raw = payload.get("actual_inbound_date")
        line.actual_inbound_date = (
            None if raw is None or str(raw).strip() == "" else _parse_date(raw, "실제입고일")
        )
        if line.actual_inbound_date is not None:
            line.actual_inbound_note = None
        touched = True
    if "actual_inbound_note" in payload:
        note_raw = str(payload.get("actual_inbound_note") or "").strip()
        note = note_raw or None
        if note and len(note) > _INBOUND_NOTE_MAX:
            raise HTTPException(
                status_code=400,
                detail=f"실제입고 비고는 {_INBOUND_NOTE_MAX}자 이내여야 합니다.",
            )
        line.actual_inbound_note = note
        if note is not None:
            line.actual_inbound_date = None
        touched = True
    if "quantity" in payload:
        line.quantity = float(_parse_decimal(payload.get("quantity"), "입고수량"))
        touched = True
    if "inbound_status" in payload:
        status = str(payload.get("inbound_status") or "").strip().upper()
        if status not in ("O", "X"):
            raise HTTPException(status_code=400, detail="입고여부는 O 또는 X 여야 합니다.")
        line.inbound_status = status
        if status == "O":
            line.expected_inbound_date = None
        touched = True
    if "line_memo" in payload:
        m = str(payload.get("line_memo") or "").strip() or None
        if m and len(m) > _LINE_MEMO_MAX:
            raise HTTPException(
                status_code=400,
                detail=f"비고 메모는 {_LINE_MEMO_MAX}자 이내여야 합니다.",
            )
        line.line_memo = m
        touched = True

    if not touched:
        raise HTTPException(status_code=400, detail="수정할 항목이 없습니다.")

    status = str(line.inbound_status or "O").strip().upper()
    ad = line.actual_inbound_date
    note = str(line.actual_inbound_note or "").strip()
    if status == "O" and ad is None and not note:
        raise HTTPException(
            status_code=400,
            detail="입고 완료(O)이면 실제입고일 또는 실제입고 비고(예: 예외 입고·무상 입고)가 필요합니다.",
        )

    db.commit()
    db.refresh(line)
    return line


def purchase_order_to_dict(po: PurchaseOrder) -> dict:
    lines = sorted(po.inbound_lines or [], key=lambda x: x.line_no)
    return {
        "id": str(po.id),
        "order_date": po.order_date.isoformat() if po.order_date else None,
        "order_date_note": po.order_date_note,
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
                "delivery_available_date": (
                    line.delivery_available_date.isoformat() if line.delivery_available_date else None
                ),
                "expected_inbound_date": (
                    None
                    if (line.inbound_status or "").strip().upper() == "O"
                    else (
                        line.expected_inbound_date.isoformat() if line.expected_inbound_date else None
                    )
                ),
                "actual_inbound_date": line.actual_inbound_date.isoformat() if line.actual_inbound_date else None,
                "actual_inbound_note": line.actual_inbound_note,
                "line_memo": line.line_memo,
                "quantity": float(line.quantity),
                "inbound_status": line.inbound_status,
                "created_at": line.created_at.isoformat() if line.created_at else None,
            }
            for line in lines
        ],
    }
