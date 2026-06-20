from pydantic import BaseModel, Field


class PurchaseInboundLineResponse(BaseModel):
    id: str
    line_no: int
    ref_code: str
    delivery_available_date: str | None = None
    expected_inbound_date: str | None = None
    actual_inbound_date: str | None = None
    actual_inbound_note: str | None = None
    line_memo: str | None = None
    quantity: float
    inbound_status: str
    created_at: str | None = None


class PurchaseOrderResponse(BaseModel):
    id: str
    order_date: str | None = None
    order_date_note: str | None = None
    erp_po_number: str
    product_type: str
    sku: str
    brand: str
    product_name: str
    manufacturer: str
    total_quantity: float
    delivery_available_date: str | None = None
    expected_inbound_date: str | None = None
    created_at: str | None = None
    inbound_lines: list[PurchaseInboundLineResponse] = Field(default_factory=list)


class PurchaseOrderListResponse(BaseModel):
    items: list[PurchaseOrderResponse] = Field(default_factory=list)


class PurchaseOrderCreateRequest(BaseModel):
    order_date: str | None = None
    order_date_note: str | None = None
    erp_po_number: str = ""
    product_type: str = "본품"
    sku: str = ""
    brand: str = ""
    product_name: str = ""
    manufacturer: str = ""
    total_quantity: float | str
    delivery_available_date: str | None = None
    expected_inbound_date: str | None = None


class PurchaseInboundCreateRequest(BaseModel):
    delivery_available_date: str | None = None
    expected_inbound_date: str | None = None
    actual_inbound_date: str | None = None
    actual_inbound_note: str | None = None
    quantity: float | str
    inbound_status: str = "O"


class PurchaseOrderInboundLineUpdate(BaseModel):
    id: str
    delivery_available_date: str | None = None
    expected_inbound_date: str | None = None
    actual_inbound_date: str | None = None
    actual_inbound_note: str | None = None
    line_memo: str | None = None
    quantity: float | str
    inbound_status: str = "O"


class PurchaseInboundLineQuickPatchRequest(BaseModel):
    delivery_available_date: str | None = None
    actual_inbound_date: str | None = None
    actual_inbound_note: str | None = None
    expected_inbound_date: str | None = None
    quantity: float | str | None = None
    inbound_status: str | None = None
    line_memo: str | None = None


class PurchaseOrderUpdateRequest(BaseModel):
    order_date: str | None = None
    order_date_note: str | None = None
    erp_po_number: str
    product_type: str = "본품"
    sku: str
    brand: str = ""
    product_name: str = ""
    manufacturer: str = ""
    total_quantity: float | str
    delivery_available_date: str | None = None
    expected_inbound_date: str | None = None
    inbound_lines: list[PurchaseOrderInboundLineUpdate] | None = None


class PurchaseOrderImportRow(BaseModel):
    row_number: int = Field(..., ge=1)
    order_date: str = ""
    erp_po_number: str = ""
    product_type: str = ""
    product_number: str = ""
    manufacturer: str = ""
    total_quantity: str | int | float = ""
    delivery_available_date: str | None = None
    expected_inbound_date: str | None = None
    actual_inbound: str | None = None
    inbound_quantity: str | int | float = ""
    inbound_status: str = ""


class PurchaseOrderImportRequest(BaseModel):
    rows: list[PurchaseOrderImportRow]


class PurchaseOrderImportResponse(BaseModel):
    created_count: int
