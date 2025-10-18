# API Dependencies for FastAPI application
# - Authentication and authorization
# - Database connection management
# - File upload validation
# - Pagination and filtering
# - Error handling helpers

from typing import Optional, Generator, Any
from datetime import datetime, UTC
from fastapi import Depends, HTTPException, status, UploadFile, Query
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from jose import JWTError, jwt
from bson import ObjectId

# Import your services and models
from core.database import get_collection, get_database
from core.security import verify_token, get_current_user_from_token
from services.user_service import get_user_by_id
from models.user import UserInDB
from models.file import FileType

# Security scheme
security = HTTPBearer()

# Database dependency
async def get_db():
    """Get database connection"""
    try:
        db = get_database()
        yield db
    finally:
        # Close connection if needed
        pass

# Authentication dependencies
async def get_current_user(
    credentials: HTTPAuthorizationCredentials = Depends(security)
) -> UserInDB:
    """Get current authenticated user from JWT token"""
    try:
        # Extract token from Authorization header
        token = credentials.credentials
        
        # Get user using the token
        user = await get_current_user_from_token(token)
        
        if user is None:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Invalid authentication credentials",
                headers={"WWW-Authenticate": "Bearer"},
            )
        
        return user
        
    except JWTError:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Could not validate credentials",
            headers={"WWW-Authenticate": "Bearer"},
        )
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Authentication failed",
            headers={"WWW-Authenticate": "Bearer"},
        )

async def get_current_active_user(
    current_user: UserInDB = Depends(get_current_user)
) -> UserInDB:
    """Get current active user (not disabled)"""
    if not current_user.is_active:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, 
            detail="Inactive user"
        )
    return current_user

# File upload dependencies
async def validate_file_upload(
    file: UploadFile,
    max_size: int = 50 * 1024 * 1024,  # 50MB
    allowed_types: list = None
) -> UploadFile:
    """Validate uploaded file size and type"""
    if allowed_types is None:
        allowed_types = [".xlsx", ".xls", ".csv"]
    
    # Check file size
    if hasattr(file, 'size') and file.size:
        if file.size > max_size:
            raise HTTPException(
                status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                detail=f"File size {file.size} exceeds maximum allowed size {max_size} bytes"
            )
    
    # Check file type by extension
    if file.filename:
        file_extension = file.filename.lower().split('.')[-1]
        if f".{file_extension}" not in allowed_types:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"File type .{file_extension} not allowed. Allowed types: {allowed_types}"
            )
    
    # Check content type
    allowed_content_types = [
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  # .xlsx
        "application/vnd.ms-excel",  # .xls
        "text/csv"  # .csv
    ]
    
    if file.content_type and file.content_type not in allowed_content_types:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Content type {file.content_type} not allowed"
        )
    
    return file

# Pagination dependencies
async def get_pagination_params(
    skip: int = Query(0, ge=0, description="Number of items to skip"),
    limit: int = Query(20, ge=1, le=100, description="Number of items to return")
) -> dict:
    """Get pagination parameters"""
    return {"skip": skip, "limit": limit}

# Permission dependencies
async def require_subscription(
    required_plan: str,
    current_user: UserInDB = Depends(get_current_active_user)
) -> UserInDB:
    """Require specific subscription level"""
    user_plan = current_user.subscription.plan
    
    # Convert enum to string if needed
    if hasattr(user_plan, 'value'):
        user_plan = user_plan.value
    
    # Define plan hierarchy (higher number = better plan)
    plan_hierarchy = {
        "free": 0,
        "basic": 1, 
        "pro": 2,
        "enterprise": 3
    }
    
    user_level = plan_hierarchy.get(user_plan, 0)
    required_level = plan_hierarchy.get(required_plan, 3)
    
    if user_level < required_level:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=f"This feature requires {required_plan} subscription or higher. Current plan: {user_plan}"
        )
    
    return current_user

def require_credits(required_credits: int):
    """Factory function to create a credits requirement dependency"""
    async def credits_dependency(
        current_user: UserInDB = Depends(get_current_active_user)
    ) -> UserInDB:
        """Require minimum credits"""
        remaining_credits = (current_user.subscription.monthly_credits_limit - 
                            current_user.subscription.monthly_credits_used)
        
        if remaining_credits < required_credits:
            raise HTTPException(
                status_code=status.HTTP_402_PAYMENT_REQUIRED,
                detail=f"Insufficient credits. Required: {required_credits}, Available: {remaining_credits}"
            )
        
        return current_user
    
    return credits_dependency