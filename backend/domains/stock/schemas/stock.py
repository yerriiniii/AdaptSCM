from typing import Any

from pydantic import BaseModel, Field


class InventoryAggregateResponse(BaseModel):
    summary: dict[str, int]
    countries: list[str]
    dates: list[str]
    rows: list[dict]
    files: list[dict] = Field(default_factory=list)
    channels: list[str] = Field(default_factory=list)
    shipment_month_columns: list[dict[str, str]] = Field(default_factory=list)
    shipment_day_labels: dict[str, str] = Field(default_factory=dict)
    shipment_totals: dict[str, Any] = Field(default_factory=dict)
    shipment_active_channel: str | None = None
    shipment_channels_with_data: list[str] = Field(default_factory=list)
    shipment_available_years: list[int] = Field(default_factory=list)
    shipment_matrix_year: int | None = None


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
    base_date: str | None = None
