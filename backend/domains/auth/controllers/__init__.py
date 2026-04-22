from domains.auth.controllers.auth_admin_api_controller import router as auth_admin_router
from domains.auth.controllers.auth_api_controller import router as auth_router

__all__ = ["auth_admin_router", "auth_router"]
