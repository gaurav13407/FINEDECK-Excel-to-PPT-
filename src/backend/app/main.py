# FastAPI Application Entry Point
# This is the main file that starts the FastAPI application

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from contextlib import asynccontextmanager
from slowapi.errors import RateLimitExceeded
from slowapi import _rate_limit_exceeded_handler
import os
import redis

from core.config import settings
from core.database import connect_to_mongo, close_mongo_connection
from core.rate_limit import limiter
from api.v1.api import api_router


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    await connect_to_mongo()
    
    # Initialize Redis for session management (optional)
    redis_client = None
    enable_sessions = os.getenv("ENABLE_REDIS_SESSIONS", "false").lower() == "true"
    
    if enable_sessions:
        redis_url = os.getenv("UPSTASH_REDIS_URL")
        if redis_url:
            try:
                redis_client = redis.from_url(redis_url)
                redis_client.ping()
                print("✅ Connected to Redis successfully!")
            except Exception as e:
                print(f"⚠️  WARNING: Failed to connect to Redis: {e}")
                redis_client = None
        else:
            print("⚠️  WARNING: ENABLE_REDIS_SESSIONS is true but UPSTASH_REDIS_URL not found")
    
    # Store Redis client and session secret in app state
    app.state.redis = redis_client
    app.state.session_secret = os.getenv("SESSION_SECRET_KEY")
    
    yield
    
    # Shutdown
    await close_mongo_connection()
    if redis_client:
        redis_client.close()


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
    expose_headers=["X-Slides-Created", "X-Template-Used", "X-AI-Features", "X-AI-Tokens", "X-AI-Cost"]  # Expose custom headers to frontend
)

# Add Session Middleware (optional - only if Redis is configured)
# This middleware will be added after app startup when Redis is initialized
# Check app.state.redis in lifespan function

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
    import os
    
    # Get port from environment (Render sets $PORT)
    port = int(os.getenv("PORT", 8000))
    
    # Disable reload in production
    reload = os.getenv("ENVIRONMENT", "development") == "development"
    
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port,
        reload=reload
    )