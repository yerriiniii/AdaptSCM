from pydantic import BaseModel, Field


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
    us_name: str | None = None
    us_sku: str | None = None
    tw_name: str | None = None
    tw_sku: str | None = None
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


class InventorySkuMappingItemResponse(BaseModel):
    group_id: str
    kr_name: str
    locales: list[InventorySkuMappingLocaleResponse] = Field(default_factory=list)


class InventorySkuMappingListResponse(BaseModel):
    items: list[InventorySkuMappingItemResponse] = Field(default_factory=list)

