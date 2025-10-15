# FastAPI Dependencies
# Common dependency injection functions used across endpoints
# - Current user authentication dependency
# - Database session dependency
# - Rate limiting dependencies
# - File upload validation dependencies
# - Admin/role-based access dependencies
# - Request validation and preprocessing
from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPBearer, HTTPAuthorizationCredentials
from typing import Optional
# Change relative imports to absolute for testing
try:
    from .database import get_database
    from .security import verify_token
except ImportError:
    # Fallback for direct execution
    from src.backend.app.core.database import get_database
    from src.backend.app.core.security import verify_token

# Security schema for extracting tokens
security = HTTPBearer()

async def get_current_user(credentials: HTTPAuthorizationCredentials = Depends(security)):
    # Extract token and validate user
    try:
        # get token from authorization header
        token=credentials.credentials

        #verify token and get user data
        payload=verify_token(token)
        if payload is None:
            raise HTTPException(
                status_code=status.HTTP_401_UNAUTHORIZED,
                detail="Invalid authentication credentials",
                headers={"WWW-Authenticate": "Bearer"}
            )
        return payload
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Could not validate credentials",
            headers={"WWW-Authenticate": "Bearer"}
        ) 

async def get_current_active_user(current_user: dict = Depends(get_current_user)):
    # check if user account is active
    return current_user

def get_db():
    # return database connection/session
    from .database import get_database
    return get_database()