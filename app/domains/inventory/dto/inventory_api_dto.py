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
    replaced_count: int = 0


class InventorySkuMappingUpsertRequest(BaseModel):
    description: str
    kr: str
    us: str | None = None
    tw: str | None = None
    vn: str | None = None
    sg: str | None = None
    au: str | None = None
    uk: str | None = None
    ae: str | None = None


class InventorySkuMappingItemResponse(BaseModel):
    description: str
    kr: str
    us: str | None = None
    tw: str | None = None
    vn: str | None = None
    sg: str | None = None
    au: str | None = None
    uk: str | None = None
    ae: str | None = None


class InventorySkuMappingListResponse(BaseModel):
    items: list[InventorySkuMappingItemResponse] = Field(default_factory=list)

