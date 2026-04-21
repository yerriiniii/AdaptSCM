import asyncio
import logging

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from sqlalchemy.exc import OperationalError

from app.domains.inventory.controllers.inventory_api_controller import (
    router as inventory_router,
)
from app.domains.inventory.models import inventory_models as _inventory_models  # noqa: F401
from app.shared.config import get_runtime_settings
from app.shared.db.session import initialize_database, is_database_configured, test_database_connection
from app.shared.storage import is_s3_configured

settings = get_runtime_settings()
app = FastAPI(title="Inventory Aggregate API", version="0.1.0")
_log = logging.getLogger(__name__)


@app.exception_handler(OperationalError)
async def database_operational_error_handler(request: Request, exc: OperationalError) -> JSONResponse:
    """DNS 실패·DB 다운 등 연결 불가 시 500 대신 503과 안내."""
    _log.warning("Database operational error on %s: %s", request.url.path, exc)
    return JSONResponse(
        status_code=503,
        content={
            "detail": "데이터베이스에 연결할 수 없습니다. PC 인터넷·DNS와 DATABASE_URL(호스트 이름)을 확인하세요. "
            "Supabase를 쓰는 경우 프로젝트가 일시 중지되었거나 호스트가 바뀌지 않았는지도 확인하세요.",
        },
    )


app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.frontend_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(inventory_router)


@app.on_event("startup")
async def on_startup() -> None:
    """요청 수락 전에 DB 스키마 1회 적용(동시 요청 경쟁 방지). 무거운 작업은 이벤트 루프를 막지 않도록 스레드에서 실행."""
    if is_database_configured():
        await asyncio.to_thread(initialize_database)
    _log.info("FastAPI startup 훅 완료")


@app.get("/health")
def health() -> dict[str, str | bool | list[str]]:
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

