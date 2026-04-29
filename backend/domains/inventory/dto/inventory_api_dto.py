import re

from typing import Any

from pydantic import BaseModel, Field, field_validator


class InventoryAggregateResponse(BaseModel):
    summary: dict[str, int]
    countries: list[str]
    dates: list[str]
    rows: list[dict]
    files: list[dict] = Field(default_factory=list)
    # 출고: 칩(시트명) 순서 또는 레거시 판매처 탭 순서
    channels: list[str] = Field(default_factory=list)
    # 출고 현황 매트릭스(월 합계 열·일자 라벨 등)
    shipment_month_columns: list[dict[str, str]] = Field(default_factory=list)
    shipment_day_labels: dict[str, str] = Field(default_factory=dict)
    shipment_totals: dict[str, Any] = Field(default_factory=dict)
    shipment_active_channel: str | None = None
    shipment_channels_with_data: list[str] = Field(default_factory=list)
    shipment_available_years: list[int] = Field(
        default_factory=list,
        description="DB에 출고 매트릭스가 있는 연도(오름차순). 파일 없으면 빈 배열.",
    )
    shipment_matrix_year: int | None = Field(
        None,
        description="현재 응답에 포함된 매트릭스 데이터 연도(shipment/view?year=).",
    )


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


class InventoryFilePatchBaseDateRequest(BaseModel):
    """재고(인벤토리) 업로드 파일의 집계 기준일. 빈 문자열·null 이면 미지정(NULL)."""

    base_date: str | None = None


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
    barcode: str | None = Field(None, max_length=255)
    option: str | None = None
    us_name: str | None = None
    us_sku: str | None = None
    tw_name: str | None = None
    tw_sku: str | None = None
    hk_name: str | None = None
    hk_sku: str | None = None
    jp_name: str | None = None
    jp_sku: str | None = None
    sg_name: str | None = None
    sg_sku: str | None = None
    de_name: str | None = None
    de_sku: str | None = None
    uk_name: str | None = None
    uk_sku: str | None = None
    au_name: str | None = None
    au_sku: str | None = None
    ae_name: str | None = None
    ae_sku: str | None = None
    vn_name: str | None = None
    vn_sku: str | None = None
    th_name: str | None = None
    th_sku: str | None = None

    @field_validator("brand")
    @classmethod
    def normalize_brand(cls, v: object) -> str:
        s = re.sub(r"\s+", " ", str(v or "").strip())[:255]
        if not s:
            raise ValueError("브랜드는 필수입니다.")
        return s


class InventorySkuMappingSegmentPatchRequest(BaseModel):
    """상품 단종 관리: 한국 상품 그룹(item) 단종 여부만 반영."""

    group_id: str = Field(..., min_length=1)
    discontinued: bool


class InventorySkuMappingItemResponse(BaseModel):
    group_id: str
    kr_name: str
    brand: str | None = None
    barcode: str | None = None
    # item.segment — 단종만 저장. 한국 상품 그룹 단줄 단종 표시용.
    segment: str | None = None
    segment_display: str | None = None
    locales: list[InventorySkuMappingLocaleResponse] = Field(default_factory=list)


class InventorySkuMappingListResponse(BaseModel):
    items: list[InventorySkuMappingItemResponse] = Field(default_factory=list)


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
    """저장된 발주 표에서 납품가능일·실제입고일·입고예정일·수량·입고여부·비고 메모만 부분 수정."""

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


class SkuLookupForPurchaseResponse(BaseModel):
    matched: bool
    brand: str = ""
    product_name: str = ""
    kr_sku: str = ""
    message: str = ""


class PurchaseOrderImportRow(BaseModel):
    """입고 엑셀 한 행 (row_number는 원본 시트의 행 번호, 에러 표시용)."""

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

