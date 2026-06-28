import json

from fastapi import APIRouter, Depends, File, Form, HTTPException, Query, UploadFile
from sqlalchemy.orm import Session

from domains.auth.dependencies import require_active_user
from domains.shipment.schemas import InventoryAggregateResponse
from domains.shipment.services.shipment_aggregate import _load_kr_sku_brand_name
from domains.shipment.services.shipment_matrix import DEFAULT_SHIPMENT_MATRIX_YEAR
from domains.shipment.services.shipment_persistence import (
    get_shipment_view,
    persist_shipment_matrix_uploads,
)
from domains.shipment.services.shipment_vendor_channel_maps import (
    parse_shipment_vendor_primary_overrides_from_payload,
)
from domains.stock.services.stock_persistence import _read_upload_bytes
from shared.branding import API_PREFIX
from shared.config import get_runtime_settings
from shared.db import get_db_session

router = APIRouter(
    prefix=f"{API_PREFIX}/shipment",
    tags=["shipment"],
    dependencies=[Depends(require_active_user)],
)


def _require_db_configured() -> None:
    if not get_runtime_settings().database_enabled:
        raise HTTPException(status_code=503, detail="DATABASE_URL is not configured.")


@router.post("/aggregate", response_model=InventoryAggregateResponse)
def shipment_aggregate(
    files: list[UploadFile] = File(...),
    shipment_vendor_primary_overrides: str | None = Form(default=None),
    db: Session = Depends(get_db_session),
) -> InventoryAggregateResponse:
    _require_db_configured()
    kr_sku_map = _load_kr_sku_brand_name(db)
    paired: list[tuple[UploadFile, bytes]] = [(uf, _read_upload_bytes(uf)) for uf in files]
    vendor_ov: dict[str, str] | None = None
    if shipment_vendor_primary_overrides and str(shipment_vendor_primary_overrides).strip():
        try:
            raw_obj = json.loads(shipment_vendor_primary_overrides)
        except json.JSONDecodeError as exc:
            raise HTTPException(
                status_code=400,
                detail="shipment_vendor_primary_overrides must be valid JSON.",
            ) from exc
        try:
            vendor_ov = parse_shipment_vendor_primary_overrides_from_payload(raw_obj)
        except ValueError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc
    persisted_files, years_uploaded = persist_shipment_matrix_uploads(
        db, paired, kr_sku_brand_name=kr_sku_map, vendor_primary_overrides_internal=vendor_ov
    )
    year_view = max(years_uploaded) if years_uploaded else DEFAULT_SHIPMENT_MATRIX_YEAR
    return InventoryAggregateResponse(**{**get_shipment_view(db, data_year=year_view), "files": persisted_files})


@router.get("/view", response_model=InventoryAggregateResponse)
def shipment_view(
    channel: str | None = Query(default=None),
    all_channels: bool = Query(default=False),
    year: int = Query(default=DEFAULT_SHIPMENT_MATRIX_YEAR, ge=2000, le=2100),
    db: Session = Depends(get_db_session),
) -> InventoryAggregateResponse:
    _require_db_configured()
    return InventoryAggregateResponse(
        **get_shipment_view(db, channel=channel, data_year=year, all_channels=all_channels)
    )
