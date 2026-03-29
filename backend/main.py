from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.domains.inventory.controllers.inventory_api_controller import (
    router as inventory_router,
)
from app.domains.inventory.models import inventory_models as _inventory_models  # noqa: F401
from app.shared.config import get_runtime_settings
from app.shared.db.session import (
    initialize_database,
    is_database_configured,
    test_database_connection,
)
from app.shared.storage import is_s3_configured

settings = get_runtime_settings()
app = FastAPI(title="Inventory Aggregate API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.frontend_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(inventory_router)


@app.on_event("startup")
def on_startup() -> None:
    if is_database_configured():
        initialize_database()


@app.get("/health")
def health() -> dict[str, str | bool]:
    database_status = "disabled"
    if is_database_configured():
        try:
            test_database_connection()
            database_status = "ok"
        except Exception:
            database_status = "error"

    return {
        "status": "ok",
        "database": database_status,
        "s3_configured": is_s3_configured(),
        "frontend_origins": settings.frontend_origins,
    }

