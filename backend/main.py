from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.domains.inventory.controllers.inventory_api_controller import (
    router as inventory_router,
)

app = FastAPI(title="Inventory Aggregate API", version="0.1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(inventory_router)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}

