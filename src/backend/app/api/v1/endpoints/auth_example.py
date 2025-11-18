"""
Authentication endpoints with Redis session management.
Example implementation showing how to use SessionManager for secure auth.
"""

from fastapi import APIRouter, Depends, HTTPException, Response, Cookie, Request
from pydantic import BaseModel
from typing import Optional
import sys
from pathlib import Path

# Add utils to path if not already there
utils_path = Path(__file__).parent.parent.parent.parent / "utils"
if str(utils_path) not in sys.path:
    sys.path.insert(0, str(utils_path.parent))

from utils.session import SessionManager


router = APIRouter(prefix="/auth", tags=["authentication"])


class LoginRequest(BaseModel):
    username: str
    password: str


class LoginResponse(BaseModel):
    message: str
    session_id: str
    user: dict


class UserResponse(BaseModel):
    user_id: str
    username: str
    email: Optional[str] = None


def get_session_manager(request: Request) -> SessionManager:
    """Dependency to get SessionManager from app state."""
    return SessionManager(request.app.state.redis)


@router.post("/login", response_model=LoginResponse)
async def login(
    login_data: LoginRequest,
    response: Response,
    session_manager: SessionManager = Depends(get_session_manager)
):
    """
    Login endpoint - validates credentials and creates a session.
    
    In production, replace the dummy check with actual database validation.
    """
    # TODO: Replace with actual database validation
    # This is a dummy check for demonstration
    if login_data.username == "admin" and login_data.password == "password":
        # Create user data
        user_data = {
            "user_id": "user_123",
            "username": login_data.username,
            "email": f"{login_data.username}@example.com",
            "roles": ["user"]
        }
        
        # Create session
        session_id = session_manager.create_session(user_data, ttl=86400)  # 24 hours
        
        # Set session cookie (httponly for security)
        response.set_cookie(
            key="session_id",
            value=session_id,
            httponly=True,
            max_age=86400,
            samesite="lax",
            secure=False  # Set to True in production with HTTPS
        )
        
        return LoginResponse(
            message="Login successful",
            session_id=session_id,
            user=user_data
        )
    
    raise HTTPException(status_code=401, detail="Invalid credentials")


@router.post("/logout")
async def logout(
    response: Response,
    session_id: Optional[str] = Cookie(None),
    session_manager: SessionManager = Depends(get_session_manager)
):
    """Logout endpoint - destroys the session."""
    if session_id:
        session_manager.delete_session(session_id)
    
    # Clear session cookie
    response.delete_cookie("session_id")
    
    return {"message": "Logged out successfully"}


@router.get("/me", response_model=UserResponse)
async def get_current_user(
    session_id: Optional[str] = Cookie(None),
    session_manager: SessionManager = Depends(get_session_manager)
):
    """
    Get current authenticated user.
    Protected endpoint that requires a valid session.
    """
    if not session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    user_data = session_manager.get_session(session_id)
    
    if not user_data:
        raise HTTPException(status_code=401, detail="Invalid or expired session")
    
    return UserResponse(**user_data)


@router.get("/protected")
async def protected_route(
    session_id: Optional[str] = Cookie(None),
    session_manager: SessionManager = Depends(get_session_manager)
):
    """
    Example protected route that requires authentication.
    """
    if not session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    user_data = session_manager.get_session(session_id)
    
    if not user_data:
        raise HTTPException(status_code=401, detail="Invalid or expired session")
    
    return {
        "message": f"Hello, {user_data.get('username')}!",
        "user": user_data
    }


@router.post("/refresh")
async def refresh_session(
    session_id: Optional[str] = Cookie(None),
    session_manager: SessionManager = Depends(get_session_manager)
):
    """Refresh session TTL."""
    if not session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    success = session_manager.refresh_session(session_id, ttl=86400)
    
    if not success:
        raise HTTPException(status_code=401, detail="Invalid or expired session")
    
    return {"message": "Session refreshed successfully"}
