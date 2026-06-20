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
    mkt_priority: str | None = None
    segment: str | None = None


class InventorySkuMappingItemResponse(BaseModel):
    group_id: str
    kr_name: str
    brand: str | None = None
    barcode: str | None = None
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

__all__ = [
    "InventorySkuMappingSummaryResponse",
    "InventorySkuMappingUploadResponse",
    "InventorySkuMappingLocaleResponse",
    "InventorySkuMappingUpsertRequest",
    "InventorySkuMappingSegmentPatchRequest",
    "InventorySkuMappingItemPatchRequest",
    "InventorySkuMappingItemResponse",
    "InventorySkuMappingListResponse",
    "SkuLookupForPurchaseResponse",
]
