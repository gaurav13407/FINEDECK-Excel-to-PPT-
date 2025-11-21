#!/usr/bin/env python3
"""
Server startup script for FinDeck FastAPI application.
This script properly sets up the Python path and starts the server with Redis session support.
"""

import os
import sys
from pathlib import Path
from dotenv import load_dotenv
import redis

# Add project root and app directory to Python path
project_root = Path(__file__).parent.parent.parent
app_dir = Path(__file__).parent / "app"
sys.path.insert(0, str(project_root))
sys.path.insert(0, str(app_dir))

# Load environment variables
load_dotenv()  

# Get Redis credentials (optional for deployments without session management)
UPSTASH_REDIS_URL = os.getenv("UPSTASH_REDIS_URL")
SESSION_SECRET_KEY = os.getenv("SESSION_SECRET_KEY")
ENABLE_REDIS_SESSIONS = os.getenv("ENABLE_REDIS_SESSIONS", "false").lower() == "true"

# Redis session setup (only if enabled and configured)
redis_client = None
session_middleware_enabled = False

if ENABLE_REDIS_SESSIONS:
    if not UPSTASH_REDIS_URL:
        print("⚠️  WARNING: ENABLE_REDIS_SESSIONS is true but UPSTASH_REDIS_URL not found")
        print("   Session middleware will be disabled. Set UPSTASH_REDIS_URL in environment variables.")
    elif not SESSION_SECRET_KEY:
        print("⚠️  WARNING: ENABLE_REDIS_SESSIONS is true but SESSION_SECRET_KEY not found")
        print("   Session middleware will be disabled. Set SESSION_SECRET_KEY in environment variables.")
        print("   Generate one with: python -c 'import secrets; print(secrets.token_hex(32))'")
    else:
        # Set up Redis client
        import redis
        try:
            redis_client = redis.from_url(UPSTASH_REDIS_URL)
            redis_client.ping()
            print("✅ Connected to Redis successfully!")
            session_middleware_enabled = True
        except Exception as e:
            print(f"⚠️  WARNING: Failed to connect to Redis: {e}")
            print("   Session middleware will be disabled.")
            redis_client = None
else:
    print("ℹ️  Redis sessions disabled (ENABLE_REDIS_SESSIONS not set to 'true')")

# Import FastAPI app (now that app_dir is in path, we can import directly from main)
from main import app

# Import middleware classes early (before conditional logic)
try:
    from middleware import SessionMiddleware
    middleware_available = True
except ImportError as e:
    print(f"⚠️  WARNING: Could not import SessionMiddleware: {e}")
    middleware_available = False

# Store Redis client in app state for use in endpoints (even if None)
app.state.redis = redis_client
app.state.session_secret = SESSION_SECRET_KEY

# ============================================================================
# MIDDLEWARE SETUP - SESSION SECURITY (Optional)
# ============================================================================
# This middleware ensures that users can't access your data by just sharing URLs
# Each user needs their own valid session to access protected routes
# Only enabled if Redis is configured

if session_middleware_enabled and redis_client and middleware_available:
    
    # Strict authentication middleware (requires session for all routes)
    app.add_middleware(
        SessionMiddleware,
        redis_client=redis_client,
        excluded_paths=[
            "/",
            "/docs",
            "/redoc",
            "/openapi.json",
            "/auth/login",
            "/auth/register",
            "/health",
            "/api/v1/auth/login",
            "/api/v1/auth/register",
            "/api/v1/users/register",
            "/api/v1/users/login",
            "/static/*",  # Exclude static files
            "/css/*",
            "/js/*",
            "/images/*",
            "/index.html",
            "/login.html",
            "/register.html",
        ],
        session_cookie_name="session_id"
    )
    print("✅ Session middleware enabled - routes are protected")
else:
    print("ℹ️  Session middleware disabled - using existing auth methods")

# ============================================================================
# SECURITY: How it works
# ============================================================================
# 1. When you login, a unique session ID is created and stored in Redis
# 2. The session ID is stored in a secure HTTP-only cookie
# 3. Every request checks if the session is valid
# 4. If someone shares your URL, they won't have your session cookie
# 5. They'll be redirected to login and can only see THEIR data
# 6. Sessions expire after 24 hours of inactivity
# ============================================================================


if __name__ == "__main__":
    import uvicorn
    
    # Get port from environment (for Render/production) or use default
    port = int(os.getenv("PORT", 8000))
    
    # Check if running in production
    is_production = os.getenv("ENVIRONMENT", "development") == "production"
    
    # Set paths for server
    backend_dir = Path(__file__).parent
    app_dir = backend_dir / "app"
    
    print("Starting FinDeck FastAPI Server...")
    print(f"Environment: {'Production' if is_production else 'Development'}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Backend directory: {backend_dir}")
    print(f"Server will be available on port: {port}")
    if not is_production:
        print(f"Server will be available at: http://localhost:{port}")
        print(f"API documentation at: http://localhost:{port}/docs")
    
    # Start server with proper module path
    uvicorn.run(
        "src.backend.app.main:app",
        host="0.0.0.0",
        port=port,
        reload=not is_production,  # Only reload in development
        reload_dirs=[str(project_root)] if not is_production else None
    )