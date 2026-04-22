"""발주 기록 CRUD."""

from __future__ import annotations

import uuid
from datetime import date
from decimal import Decimal

from fastapi import HTTPException
from sqlalchemy import func, select
from sqlalchemy.orm import Session, selectinload

from domains.inventory.models.inventory_models import PurchaseInboundLine, PurchaseOrder
from domains.inventory.services.inventory_mapping_service import lookup_product_by_any_country_sku

ORDER_DATE_PLANNED_CANONICAL = "발주 예정"
ORDER_DATE_NOTE_MAX = 128
NO_ERP_PO_PREFIX = "__NO_ERP__"
NO_SKU_PLACEHOLDER_PREFIX = "__NO_SKU__"


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


def _parse_optional_sheet_date(value: object, field_label: str) -> tuple[date | None, str | None]:
    """엑셀에서 온 날짜 문자열. 빈 값은 (None, None). 비어 있지 않은데 파싱 실패 시 에러 문구."""
    if value is None or value == "":
        return None, None
    if isinstance(value, date):
        return value, None
    s = str(value).strip()
    if not s:
        return None, None
    s2 = s.replace(".", "-").replace("/", "-")
    for cand in (s[:10], s2[:10], s, s2):
        cand = cand.strip()
        if not cand:
            continue
        try:
            parts = cand.split("-")
            if len(parts) == 3:
                y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                return date(y, m, d), None
        except (ValueError, TypeError):
            pass
        try:
            return date.fromisoformat(cand[:10]), None
        except ValueError:
            continue
    return None, f"{field_label} 날짜 형식이 올바르지 않습니다. (YYYY-MM-DD)"


def _parse_optional_sheet_date_loose(value: object) -> date | None:
    """납품가능일·입고예정일: 날짜로 읽히면 저장, 그 외(무상 입고 등 텍스트)는 None — 저장된 발주 인라인과 동일하게 날짜만 반영."""
    if value is None or value == "":
        return None
    if isinstance(value, date):
        return value
    s = str(value).strip()
    if not s:
        return None
    d, _err = _parse_optional_sheet_date(s, "_")
    return d


_TOTAL_Q_OMIT_NORMALIZED = frozenset(
    {"", "-", "–", "—", "\u2013", "\u2014", "n/a", "na", "#n/a", "없음", "해당없음", "해당 없음"}
)


def _parse_total_quantity_optional(value: object) -> tuple[Decimal | None, bool, str | None]:
    """(총 발주수량, 헤더 생략 여부, 에러). 공백·대시 등이면 추가 입고만(기존 발주 필요)."""
    if value is None:
        return None, True, None
    s_raw = str(value).replace(",", "").strip()
    if not s_raw:
        return None, True, None
    key = s_raw.casefold().replace(" ", "").replace("\u3000", "")
    if s_raw in _TOTAL_Q_OMIT_NORMALIZED or key in {"-", "–", "—", "n/a", "na", "#n/a", "해당없음"}:
        return None, True, None
    try:
        d = Decimal(s_raw)
    except Exception:
        return None, False, "총 발주수량 숫자 형식이 올바르지 않습니다."
    if d <= 0:
        return None, False, "총 발주수량은 0보다 커야 합니다."
    return d, False, None


def _normalize_inbound_status_sheet(raw: object) -> str | None:
    s = str(raw or "").strip()
    if not s:
        return None
    compact = "".join(s.split()).replace("\u3000", "").casefold()
    u = s.upper()
    if u == "O" or compact in ("입고완료", "완료"):
        return "O"
    if u == "X" or compact in ("입고예정", "예정"):
        return "X"
    if "예정" in s:
        return "X"
    if "완료" in s:
        return "O"
    return None


def _try_decimal_sheet(value: object) -> tuple[Decimal | None, str | None]:
    if value is None:
        return None, "값이 비어 있습니다."
    s = str(value).replace(",", "").strip()
    if not s:
        return None, "값이 비어 있습니다."
    try:
        return Decimal(s), None
    except Exception:
        return None, "숫자 형식이 올바르지 않습니다."


def _actual_inbound_from_sheet(raw: object) -> tuple[date | None, str | None, str | None]:
    """(날짜, 비고, 에러). 저장된 발주 인라인과 동일: YYYY-MM-DD면 날짜, 아니면 비고 텍스트(무상 입고 등)."""
    s = str(raw or "").strip()
    if not s:
        return None, None, None
    d, _err = _parse_optional_sheet_date(s, "실제입고일")
    if d is not None:
        return d, None, None
    if len(s) > _INBOUND_NOTE_MAX:
        return None, None, f"실제입고 비고는 {_INBOUND_NOTE_MAX}자 이내여야 합니다."
    return None, s, None


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


def create_purchase_order(db: Session, payload: dict, *, commit: bool = True) -> PurchaseOrder:
    erp = str(payload.get("erp_po_number") or "").strip()
    if not erp:
        erp = f"{NO_ERP_PO_PREFIX}{uuid.uuid4().hex}"

    sku = str(payload.get("sku") or "").strip()
    if not sku:
        sku = f"{NO_SKU_PLACEHOLDER_PREFIX}{uuid.uuid4().hex}"

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
    db.flush()
    db.refresh(po)

    # 저장 표는 입고 차수 행이 있어야 실제입고·수량·입고여부·비고 연필이 보인다. 신규 발주에 1차 입고를 둔다.
    line = PurchaseInboundLine(
        purchase_order_id=po.id,
        line_no=1,
        ref_code="_",
        delivery_available_date=po.delivery_available_date,
        expected_inbound_date=po.expected_inbound_date,
        actual_inbound_date=None,
        actual_inbound_note=None,
        quantity=float(total_q),
        inbound_status="X",
    )
    db.add(line)
    db.flush()
    po_loaded = get_purchase_order(db, po.id)
    if po_loaded is not None:
        _normalize_inbound_ref_codes(po_loaded)

    if commit:
        db.commit()
    else:
        db.flush()
    db.refresh(po)
    return po


def add_inbound_line(
    db: Session, order_id: uuid.UUID, payload: dict, *, commit: bool = True
) -> PurchaseInboundLine:
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
    if commit:
        db.commit()
    else:
        db.flush()
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


def patch_inbound_line_quick(
    db: Session, order_id: uuid.UUID, line_id: uuid.UUID, payload: dict, *, commit: bool = True
) -> PurchaseInboundLine:
    """실제입고일·입고수량·입고여부·비고 메모만 부분 수정."""
    po = get_purchase_order(db, order_id)
    if po is None:
        raise HTTPException(status_code=404, detail="발주를 찾을 수 없습니다.")
    line = next((ln for ln in (po.inbound_lines or []) if ln.id == line_id), None)
    if line is None:
        raise HTTPException(status_code=404, detail="입고 차수를 찾을 수 없습니다.")

    touched = False
    if "delivery_available_date" in payload:
        raw_d = payload.get("delivery_available_date")
        line.delivery_available_date = (
            None if raw_d is None or str(raw_d).strip() == "" else _parse_date(raw_d, "납품가능일")
        )
        touched = True
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
    if "expected_inbound_date" in payload:
        status_now = str(line.inbound_status or "").strip().upper()
        raw_e = payload.get("expected_inbound_date")
        clearing = raw_e is None or str(raw_e).strip() == ""
        if clearing:
            if status_now != "O":
                line.expected_inbound_date = None
            touched = True
        elif status_now == "O":
            raise HTTPException(
                status_code=400,
                detail="입고 완료(O) 차수는 입고예정일을 수정할 수 없습니다.",
            )
        else:
            line.expected_inbound_date = _parse_date(raw_e, "입고예정일")
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

    if commit:
        db.commit()
    else:
        db.flush()
    db.refresh(line)
    return line


def import_purchase_orders_from_sheet(db: Session, rows: list[dict]) -> int:
    """입고 엑셀에서 파싱된 행들을 검증한 뒤 저장된 발주로 일괄 저장. 전부 성공하거나 전부 롤백.

    동일 ERP PO·상품코드(SKU)·발주일이면 하나의 발주로 보고, 첫 행은 발주 생성+1차 입고 반영,
    이후 행(또는 DB에 이미 있는 발주)은 입고 차수만 추가한다.
    """
    errors: list[str] = []
    prepared: list[dict] = []

    for item in rows:
        rn = int(item.get("row_number") or 0)
        erp = str(item.get("erp_po_number") or "").strip()
        order_date_raw = str(item.get("order_date") or "").strip()
        pnum = str(item.get("product_number") or "").strip()
        if not erp and not pnum and not order_date_raw:
            continue
        row_errs: list[str] = []
        if not erp:
            row_errs.append(f"{rn}행: ERP PO번호가 비어 있습니다.")
        if not order_date_raw:
            row_errs.append(f"{rn}행: 발주일자가 비어 있습니다.")
        else:
            od, od_err = _parse_optional_sheet_date(order_date_raw, "발주일자")
            if od_err:
                row_errs.append(f"{rn}행: 발주일자 — {od_err}")
            elif od is None:
                row_errs.append(f"{rn}행: 발주일자가 비어 있습니다.")
        if not pnum:
            row_errs.append(f"{rn}행: 상품번호가 비어 있습니다.")
        total_d, total_omit, t_err = _parse_total_quantity_optional(item.get("total_quantity"))
        if t_err:
            row_errs.append(f"{rn}행: 총 발주수량 — {t_err}")
        in_q, in_err = _try_decimal_sheet(item.get("inbound_quantity"))
        if in_err:
            row_errs.append(f"{rn}행: 입고수량 — {in_err}")
        elif in_q is not None and in_q < 0:
            row_errs.append(f"{rn}행: 입고수량은 0 이상이어야 합니다.")

        st = _normalize_inbound_status_sheet(item.get("inbound_status"))
        if st is None:
            row_errs.append(
                f"{rn}행: 입고여부를 인식할 수 없습니다. "
                f"(입고완료·완료·O 또는 입고 예정·예정·X)"
            )

        deliv = _parse_optional_sheet_date_loose(item.get("delivery_available_date"))
        exp = _parse_optional_sheet_date_loose(item.get("expected_inbound_date"))

        ad: date | None = None
        an: str | None = None
        act_err: str | None = None
        if item.get("actual_inbound") is not None and str(item.get("actual_inbound") or "").strip():
            ad, an, act_err = _actual_inbound_from_sheet(item.get("actual_inbound"))
            if act_err:
                row_errs.append(f"{rn}행: {act_err}")

        if st == "O" and ad is None and not (an and str(an).strip()):
            row_errs.append(
                f"{rn}행: 입고여부가 완료(O)이면 실제입고일(YYYY-MM-DD) 또는 비고 텍스트가 필요합니다."
            )

        if row_errs:
            errors.extend(row_errs)
            continue

        if st is None or in_q is None:
            errors.append(f"{rn}행: 필수 값을 확인해 주세요.")
            continue
        if not total_omit and total_d is None:
            errors.append(f"{rn}행: 총 발주수량을 입력하세요.")
            continue
        od2, _od2e = _parse_optional_sheet_date(order_date_raw, "발주일자")
        if od2 is None:
            errors.append(f"{rn}행: 발주일자를 확인해 주세요.")
            continue

        hit = lookup_product_by_any_country_sku(db, pnum)
        if not hit.get("matched"):
            errors.append(f"{rn}행: 상품번호 — {hit.get('message') or '매칭 실패'}")
            continue

        kr_sku = str(hit.get("kr_sku") or "").strip()
        if not kr_sku:
            errors.append(f"{rn}행: 상품번호에 대응하는 한국 SKU를 찾을 수 없습니다.")
            continue

        biz_key = (erp, kr_sku, od2)

        product_type = str(item.get("product_type") or "").strip() or "본품"
        manufacturer = str(item.get("manufacturer") or "").strip()

        patch: dict = {
            "delivery_available_date": deliv.isoformat() if deliv else None,
            "quantity": float(in_q),
            "inbound_status": st,
            "expected_inbound_date": (exp.isoformat() if exp else None) if st == "X" else None,
            "actual_inbound_date": (ad.isoformat() if ad else "") if st == "O" else "",
            "actual_inbound_note": (an if an else "") if st == "O" else "",
        }

        create_payload: dict | None = None
        if not total_omit:
            assert total_d is not None
            create_payload = {
                "order_date": od2.isoformat(),
                "order_date_note": None,
                "erp_po_number": erp,
                "product_type": product_type,
                "sku": kr_sku,
                "brand": str(hit.get("brand") or "").strip(),
                "product_name": str(hit.get("product_name") or "").strip(),
                "manufacturer": manufacturer,
                "total_quantity": float(total_d),
                "delivery_available_date": deliv.isoformat() if deliv else None,
                "expected_inbound_date": exp.isoformat() if exp else None,
            }

        prepared.append(
            {
                "row_number": rn,
                "biz_key": biz_key,
                "erp": erp,
                "kr_sku": kr_sku,
                "od2": od2,
                "total_omit": total_omit,
                "create": create_payload,
                "patch": patch,
            }
        )

    if errors:
        raise HTTPException(status_code=400, detail=errors)

    po_by_key: dict[tuple[str, str, date | None], uuid.UUID] = {}

    try:
        for pack in prepared:
            bk = pack["biz_key"]
            po_id = po_by_key.get(bk)
            if po_id is None:
                existing = find_existing_purchase_order_id_by_business_key(
                    db, pack["erp"], pack["kr_sku"], pack["od2"]
                )
                if existing is None:
                    if pack["total_omit"]:
                        raise HTTPException(
                            status_code=400,
                            detail=(
                                f'{pack["row_number"]}행: 총 발주수량이 비어 있으면 동일 ERP·SKU·발주일의 발주가 '
                                f"이미 있어야 합니다. 먼저 총 발주수량이 있는 행을 넣거나, 저장된 발주를 확인하세요."
                            ),
                        )
                    assert pack["create"] is not None
                    po = create_purchase_order(db, pack["create"], commit=False)
                    po_by_key[bk] = po.id
                    po_full = get_purchase_order(db, po.id)
                    if po_full is None or not po_full.inbound_lines:
                        raise HTTPException(
                            status_code=500,
                            detail="발주 생성 후 입고 행을 찾을 수 없습니다.",
                        )
                    first = sorted(po_full.inbound_lines, key=lambda x: x.line_no)[0]
                    patch_inbound_line_quick(db, po.id, first.id, pack["patch"], commit=False)
                else:
                    po_by_key[bk] = existing
                    add_inbound_line(db, existing, pack["patch"], commit=False)
            else:
                add_inbound_line(db, po_id, pack["patch"], commit=False)
        db.commit()
    except HTTPException:
        db.rollback()
        raise
    except Exception as exc:
        db.rollback()
        raise HTTPException(status_code=500, detail=f"저장 중 오류: {exc}") from exc

    return len(prepared)


def delete_inbound_line(db: Session, order_id: uuid.UUID, line_id: uuid.UUID) -> None:
    po = get_purchase_order(db, order_id)
    if po is None:
        raise HTTPException(status_code=404, detail="발주를 찾을 수 없습니다.")
    line = next((ln for ln in (po.inbound_lines or []) if ln.id == line_id), None)
    if line is None:
        raise HTTPException(status_code=404, detail="입고 차수를 찾을 수 없습니다.")
    db.delete(line)
    db.flush()
    stmt = (
        select(PurchaseInboundLine)
        .where(PurchaseInboundLine.purchase_order_id == order_id)
        .order_by(PurchaseInboundLine.line_no)
    )
    remaining = list(db.scalars(stmt).unique().all())
    for i, ln in enumerate(remaining, start=1):
        ln.line_no = i
    po2 = get_purchase_order(db, order_id)
    if po2 is not None:
        _normalize_inbound_ref_codes(po2)
    db.commit()


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
