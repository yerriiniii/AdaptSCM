from .admin import router as auth_admin_router
from .auth import router as auth_router

__all__ = ["auth_admin_router", "auth_router"]
