"""
Redis-based session middleware for FastAPI.
Automatically validates sessions on every request.
"""

from fastapi import Request, HTTPException
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import JSONResponse
from typing import Optional, List
from utils.session import SessionManager


class SessionMiddleware(BaseHTTPMiddleware):
    """
    Middleware that automatically validates sessions for all requests.
    Excludes certain paths like /docs, /login, etc.
    """
    
    def __init__(
        self,
        app,
        redis_client,
        excluded_paths: Optional[List[str]] = None,
        session_cookie_name: str = "session_id"
    ):
        """
        Initialize the session middleware.
        
        Args:
            app: FastAPI application
            redis_client: Redis client instance
            excluded_paths: Paths that don't require authentication
            session_cookie_name: Name of the session cookie
        """
        super().__init__(app)
        self.session_manager = SessionManager(redis_client)
        self.session_cookie_name = session_cookie_name
        
        # Default excluded paths (public endpoints)
        self.excluded_paths = excluded_paths or [
            "/",
            "/docs",
            "/redoc",
            "/openapi.json",
            "/auth/login",
            "/auth/register",
            "/health",
            "/api/v1/auth/login",
            "/api/v1/auth/register",
        ]
    
    async def dispatch(self, request: Request, call_next):
        """Process each request and validate session."""
        
        # Check if path is excluded from authentication
        if self._is_excluded_path(request.url.path):
            return await call_next(request)
        
        # Get session ID from cookie
        session_id = request.cookies.get(self.session_cookie_name)
        
        if not session_id:
            return JSONResponse(
                status_code=401,
                content={"detail": "Not authenticated. Please login."}
            )
        
        # Validate session
        user_data = self.session_manager.get_session(session_id)
        
        if not user_data:
            return JSONResponse(
                status_code=401,
                content={"detail": "Invalid or expired session. Please login again."}
            )
        
        # Add user data to request state (accessible in route handlers)
        request.state.user = user_data
        request.state.session_id = session_id
        
        # Continue to the route handler
        response = await call_next(request)
        return response
    
    def _is_excluded_path(self, path: str) -> bool:
        """Check if the path is excluded from authentication."""
        # Exact match
        if path in self.excluded_paths:
            return True
        
        # Pattern matching (e.g., /static/*)
        for excluded in self.excluded_paths:
            if excluded.endswith("*") and path.startswith(excluded[:-1]):
                return True
        
        return False


class OptionalSessionMiddleware(BaseHTTPMiddleware):
    """
    Middleware that loads session data if available, but doesn't require it.
    Useful for routes that work differently for authenticated vs unauthenticated users.
    """
    
    def __init__(
        self,
        app,
        redis_client,
        session_cookie_name: str = "session_id"
    ):
        super().__init__(app)
        self.session_manager = SessionManager(redis_client)
        self.session_cookie_name = session_cookie_name
    
    async def dispatch(self, request: Request, call_next):
        """Process request and load session if available."""
        
        # Get session ID from cookie
        session_id = request.cookies.get(self.session_cookie_name)
        
        # Load session data if available (but don't require it)
        if session_id:
            user_data = self.session_manager.get_session(session_id)
            if user_data:
                request.state.user = user_data
                request.state.session_id = session_id
                request.state.is_authenticated = True
            else:
                request.state.is_authenticated = False
        else:
            request.state.is_authenticated = False
        
        response = await call_next(request)
        return response
