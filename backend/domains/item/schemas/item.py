import re

from pydantic import BaseModel, Field, field_validator


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
    skipped_overseas_count: int = 0
    skipped_overseas_warnings: list[str] = Field(default_factory=list)


class InventorySkuMappingLocaleResponse(BaseModel):
    country_code: str
    sku_type: str | None = None
    name: str | None = None
    sku: str


class InventorySkuMappingLocaleRequest(BaseModel):
    country_code: str
    sku_type: str | None = None
    sku: str | None = None
    name: str | None = None


class InventorySkuMappingUpsertRequest(BaseModel):
    kr_name: str
    kr_sku: str
    brand: str = Field(..., max_length=255)
    representative_code: str | None = Field(None, max_length=255)
    barcode: str | None = Field(None, max_length=255)
    category: str | None = Field(None, max_length=255)
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
    overseas_locales: list[InventorySkuMappingLocaleRequest] = Field(default_factory=list)
    segment: str | None = Field(None, max_length=255)
    mkt_priority: str | None = Field(None, max_length=255)

    @field_validator("brand")
    @classmethod
    def normalize_brand(cls, v: object) -> str:
        s = re.sub(r"\s+", " ", str(v or "").strip())[:255]
        if not s:
            raise ValueError("brand is required")
        return s


class InventorySkuMappingSegmentPatchRequest(BaseModel):
    group_id: str = Field(..., min_length=1)
    discontinued: bool


class InventorySkuMappingItemPatchRequest(BaseModel):
    group_id: str = Field(..., min_length=1)
    kr_sku: str = Field(..., min_length=1)
    kr_name: str = Field(..., min_length=1)
    brand: str = Field(..., min_length=1)
    barcode: str | None = Field(None, max_length=255)
    category: str | None = Field(None, max_length=255)
    mkt_priority: str | None = None
    segment: str | None = None


class InventorySkuMappingItemResponse(BaseModel):
    group_id: str
    kr_name: str
    brand: str | None = None
    representative_code: str | None = None
    barcode: str | None = None
    category: str | None = None
    segment: str | None = None
    segment_display: str | None = None
    mkt_priority: str | None = None
    locales: list[InventorySkuMappingLocaleResponse] = Field(default_factory=list)


class InventorySkuMappingListResponse(BaseModel):
    items: list[InventorySkuMappingItemResponse] = Field(default_factory=list)


class SkuLookupForPurchaseResponse(BaseModel):
    matched: bool
    brand: str = ""
    product_name: str = ""
    kr_sku: str = ""
    message: str = ""


class ItemMasterExtraColumnResponse(BaseModel):
    field_key: str
    label: str
    sort_order: int = 0


class ItemMasterExtraColumnCreateRequest(BaseModel):
    label: str = Field(..., min_length=1, max_length=64)


class ItemMasterRowResponse(BaseModel):
    group_id: str
    brand: str = ""
    representative_code: str = ""
    segment: str = ""
    version: str = ""
    kr_sku: str = ""
    kr_name: str = ""
    stock_category: str = ""
    category: str = ""
    fcst_grade: str = ""
    stock_grade: str = ""
    release_month: str = ""
    code_registered_at: str | None = None
    kr_grade: str = ""
    us_grade: str = ""
    tw_grade: str = ""
    hk_grade: str = ""
    jp_grade: str = ""
    us_codes: str = ""
    tw_codes: str = ""
    hk_codes: str = ""
    jp_codes: str = ""
    sg_codes: str = ""
    my_codes: str = ""
    de_codes: str = ""
    uk_codes: str = ""
    au_codes: str = ""
    ae_codes: str = ""
    vn_codes: str = ""
    th_codes: str = ""
    barcode: str = ""
    locales: list[InventorySkuMappingLocaleResponse] = Field(default_factory=list)
    extra_fields: dict[str, str] = Field(default_factory=dict)
    updated_at: str | None = None


class ItemMasterRowListResponse(BaseModel):
    items: list[ItemMasterRowResponse] = Field(default_factory=list)
    columns: list[str] = Field(default_factory=list)
    extra_columns: list[ItemMasterExtraColumnResponse] = Field(default_factory=list)


class ItemMasterRowDetailResponse(BaseModel):
    row: ItemMasterRowResponse
    extra_columns: list[ItemMasterExtraColumnResponse] = Field(default_factory=list)


class ItemMasterUploadResponse(BaseModel):
    processed_file_count: int = 0
    processed_row_count: int = 0
    updated_count: int = 0
    created_count: int = 0
    columns: list[str] = Field(default_factory=list)


class ItemMasterRowPatchRequest(BaseModel):
    group_id: str = Field(..., min_length=1)
    brand: str = Field(..., min_length=1)
    representative_code: str | None = None
    segment: str | None = None
    version: str | None = None
    kr_sku: str = Field(..., min_length=1)
    kr_name: str = Field(..., min_length=1)
    stock_category: str | None = None
    category: str | None = None
    fcst_grade: str | None = None
    stock_grade: str | None = None
    release_month: str | None = None
    code_registered_at: str | None = None
    kr_grade: str | None = None
    us_grade: str | None = None
    tw_grade: str | None = None
    hk_grade: str | None = None
    jp_grade: str | None = None
    barcode: str | None = Field(None, max_length=255)
    overseas_locales: list[InventorySkuMappingLocaleRequest] | None = None
    extra_fields: dict[str, str | None] | None = None

    @field_validator("brand")
    @classmethod
    def normalize_brand(cls, v: object) -> str:
        s = re.sub(r"\s+", " ", str(v or "").strip())[:255]
        if not s:
            raise ValueError("brand is required")
        return s

__all__ = [
    "InventorySkuMappingSummaryResponse",
    "InventorySkuMappingUploadResponse",
    "InventorySkuMappingLocaleResponse",
    "InventorySkuMappingLocaleRequest",
    "InventorySkuMappingUpsertRequest",
    "InventorySkuMappingSegmentPatchRequest",
    "InventorySkuMappingItemPatchRequest",
    "InventorySkuMappingItemResponse",
    "InventorySkuMappingListResponse",
    "SkuLookupForPurchaseResponse",
    "ItemMasterRowResponse",
    "ItemMasterRowListResponse",
    "ItemMasterExtraColumnResponse",
    "ItemMasterExtraColumnCreateRequest",
    "ItemMasterUploadResponse",
    "ItemMasterRowPatchRequest",
]
