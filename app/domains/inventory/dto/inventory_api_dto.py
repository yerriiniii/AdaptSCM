import re

from pydantic import BaseModel, Field, field_validator


class InventoryAggregateResponse(BaseModel):
    summary: dict[str, int]
    countries: list[str]
    dates: list[str]
    rows: list[dict]
    files: list[dict] = Field(default_factory=list)


class InventoryDirectUploadFileRequest(BaseModel):
    client_id: str
    original_name: str
    country_code: str
    date: str | None = None
    content_type: str | None = None
    size: int | None = None


class InventoryDirectUploadRequest(BaseModel):
    files: list[InventoryDirectUploadFileRequest]


class InventoryDirectUploadPreparedFile(BaseModel):
    client_id: str
    original_name: str
    country_code: str
    date: str | None = None
    stored_name: str
    s3_key: str
    upload_url: str
    file_id: str | None = None


class InventoryDirectUploadPreparedResponse(BaseModel):
    files: list[InventoryDirectUploadPreparedFile] = Field(default_factory=list)


class InventoryCompleteUploadFileRequest(BaseModel):
    client_id: str
    original_name: str
    country_code: str
    date: str | None = None
    stored_name: str
    s3_key: str
    size: int | None = None
    content_type: str | None = None


class InventoryCompleteUploadRequest(BaseModel):
    files: list[InventoryCompleteUploadFileRequest]


class InventorySkuMappingSummaryResponse(BaseModel):
    total_count: int
    updated_at: str | None = None
    upload_updated_at: str | None = None
    manual_updated_at: str | None = None
    required_columns: list[str] = Field(default_factory=list)
    optional_columns: list[str] = Field(default_factory=list)


class InventorySkuMappingUploadResponse(InventorySkuMappingSummaryResponse):
    processed_file_count: int = 0
    merged_item_count: int = 0


class InventorySkuMappingLocaleResponse(BaseModel):
    country_code: str
    name: str | None = None
    sku: str


class InventorySkuMappingUpsertRequest(BaseModel):
    kr_name: str
    kr_sku: str
    brand: str = Field(..., max_length=255)
    option: str | None = None
    us_name: str | None = None
    us_sku: str | None = None
    tw_name: str | None = None
    tw_sku: str | None = None
    hk_name: str | None = None
    hk_sku: str | None = None
    vn_name: str | None = None
    vn_sku: str | None = None
    sg_name: str | None = None
    sg_sku: str | None = None
    au_name: str | None = None
    au_sku: str | None = None
    uk_name: str | None = None
    uk_sku: str | None = None
    ae_name: str | None = None
    ae_sku: str | None = None

    @field_validator("brand")
    @classmethod
    def normalize_brand(cls, v: object) -> str:
        s = re.sub(r"\s+", " ", str(v or "").strip())[:255]
        if not s:
            raise ValueError("브랜드는 필수입니다.")
        return s


class InventorySkuMappingItemResponse(BaseModel):
    group_id: str
    kr_name: str
    brand: str | None = None
    locales: list[InventorySkuMappingLocaleResponse] = Field(default_factory=list)


class InventorySkuMappingListResponse(BaseModel):
    items: list[InventorySkuMappingItemResponse] = Field(default_factory=list)


class PurchaseInboundLineResponse(BaseModel):
    id: str
    line_no: int
    ref_code: str
    actual_inbound_date: str | None = None
    quantity: float
    inbound_status: str
    created_at: str | None = None


class PurchaseOrderResponse(BaseModel):
    id: str
    order_date: str | None = None
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
    order_date: str
    erp_po_number: str
    product_type: str = "본품"
    sku: str
    brand: str = ""
    product_name: str = ""
    manufacturer: str = ""
    total_quantity: float | str
    delivery_available_date: str | None = None
    expected_inbound_date: str | None = None


class PurchaseInboundCreateRequest(BaseModel):
    actual_inbound_date: str | None = None
    quantity: float | str
    inbound_status: str = "O"


class SkuLookupForPurchaseResponse(BaseModel):
    matched: bool
    brand: str = ""
    product_name: str = ""
    message: str = ""

