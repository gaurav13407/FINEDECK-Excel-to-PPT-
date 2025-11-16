# API Version 1 Router
# Main router that includes all API v1 endpoints
# - Combines all endpoint routers (auth, users, files, conversions)
# - Sets up common prefixes and tags for OpenAPI documentation
# - Handles API versioning and backwards compatibility
# - Global error handling for API responses# API Version 1 Router
# Main router that includes all API v1 endpoints

from fastapi import APIRouter

# Import all v1 endpoint modules (including codes and templates)
from api.v1.endpoints import auth, users, files, conversions, codes, templates, tiered_conversions, plan_upgrades
from api.v1.endpoints import oauth, intelligent_conversion, powerbi

api_router = APIRouter()

# Include all endpoint routers
api_router.include_router(
    auth.router, 
    prefix="/auth", 
    tags=["authentication"]
)

api_router.include_router(
    users.router, 
    prefix="/users", 
    tags=["users"]
)

api_router.include_router(
    files.router, 
    prefix="/files", 
    tags=["files"]
)

api_router.include_router(
    conversions.router, 
    prefix="/conversions", 
    tags=["conversions"]
)
api_router.include_router(
    codes.router,
    prefix="/codes",
    tags=["codes"]
)
api_router.include_router(
    oauth.router,
    prefix="/auth",
    tags=["oauth"]
)
api_router.include_router(
    templates.router,
    prefix="/templates",
    tags=["templates"]
)

api_router.include_router(
    tiered_conversions.router,
    prefix="/tiered",
    tags=["tiered-conversions"]
)

api_router.include_router(
    plan_upgrades.router,
    prefix="/upgrades",
    tags=["plan-upgrades"]
)

api_router.include_router(
    intelligent_conversion.router,
    prefix="/intelligence",
    tags=["intelligent-conversion"]
)

api_router.include_router(
    powerbi.router,
    prefix="/powerbi",
    tags=["power-bi-dashboards"]
)