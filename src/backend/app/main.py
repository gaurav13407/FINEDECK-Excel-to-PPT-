# FastAPI Application Entry Point
# This is the main file that starts the FastAPI application

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager
from slowapi.errors import RateLimitExceeded
from slowapi import _rate_limit_exceeded_handler

from core.config import settings
from core.database import connect_to_mongo, close_mongo_connection
from core.rate_limit import limiter
from api.v1.api import api_router


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    await connect_to_mongo()
    yield
    # Shutdown
    await close_mongo_connection()


# Create FastAPI application
app = FastAPI(
    title=settings.app_name,
    description="""
    ## FinDeck Excel to PowerPoint Conversion API
    
    Convert Excel financial data into professional PowerPoint presentations.
    
    ### Features
    * 📊 Excel file upload and parsing
    * 📈 Automated chart and table generation
    * 🎨 Professional PPT templates
    * 💳 Subscription-based credit system
    * 🔐 JWT authentication with OAuth support
    
    ### Subscription Tiers
    * **Free**: 1 credit/month, 1 file limit
    * **Basic**: 10 credits/month, 5 files
    * **Pro**: 50 credits/month, 15 files
    * **AI (Enterprise)**: 1000 credits/month, 100 files
    
    ### Authentication
    Use JWT tokens in the Authorization header:
    ```
    Authorization: Bearer <your_token>
    ```
    """,
    version="1.0.0",
    terms_of_service="https://www.findeck.live/terms",
    contact={
        "name": "FinDeck Support",
        "email": "support@findeck.live",
        "url": "https://www.findeck.live"
    },
    license_info={
        "name": "Proprietary",
        "url": "https://www.findeck.live/license"
    },
    openapi_url=f"/api/{settings.api_version}/openapi.json",
    docs_url="/api/docs",
    redoc_url="/api/redoc",
    lifespan=lifespan
)

# Set up CORS middleware with production-safe origins
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,  # Now using settings from .env
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Register rate limiter
app.state.limiter = limiter
app.add_exception_handler(RateLimitExceeded, _rate_limit_exceeded_handler)

# Include API router
app.include_router(api_router, prefix=f"/api/{settings.api_version}")


@app.get("/")
async def root():
    return {
        "message": "FinDeck Excel to PowerPoint API",
        "version": "1.0.0",
        "docs": "/docs"
    }


@app.get("/health")
async def health_check():
    return {"status": "healthy"}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True
    )