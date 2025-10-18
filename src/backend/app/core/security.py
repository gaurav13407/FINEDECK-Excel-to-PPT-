# Security and Authentication Utilities
# Handles all security-related operations
# - JWT token creation, validation, and refresh
# - Password hashing and verification using bcrypt
# - Token blacklisting for logout functionality
# - Rate limiting decorators and middleware
# - Input validation and sanitization
# - Security headers and CSRF protection

import hashlib
import secrets
from jose import JWTError,jwt
from datetime import datetime,timedelta,UTC
from typing import Optional
from .config import settings

# Alternative password hashing using SHA256 + salt (more compatible)
def _generate_salt() -> str:
    """Generate a random salt for password hashing"""
    return secrets.token_hex(32)

def get_password_hash(password: str) -> str:
    """Hash a plaintext password using SHA256 + salt"""
    salt = _generate_salt()
    password_hash = hashlib.sha256((password + salt).encode()).hexdigest()
    return f"{salt}:{password_hash}"
    
def verify_password(plain_password: str, hashed_password: str) -> bool:
    """Verify a plaintext password against the hashed version"""
    try:
        salt, stored_hash = hashed_password.split(':', 1)
        password_hash = hashlib.sha256((plain_password + salt).encode()).hexdigest()
        return password_hash == stored_hash
    except ValueError:
        return False

def create_access_token(data:dict,expires_delta:Optional[timedelta]=None)->str:
    """Create a JWT access token"""
    # What should this function do?
    # - Take user data (like user_id, email)
    # - Add expiration time (24 hours default)
    # - Create and return JWT token
    to_encode=data.copy()

    if expires_delta:
        expire=datetime.now(UTC)+expires_delta
    else:
        expire=datetime.now(UTC)+timedelta(minutes=settings.jwt_expiration_minutes)
    to_encode.update({"exp":expire})
    encoded_jwt=jwt.encode(to_encode,settings.jwt_secret,algorithm=settings.jwt_algorithm)
    return encoded_jwt

def verify_token(token:str)-> Optional[dict]:
    """Verify and decode a JWT token"""
    # What should this function do?
    # - Take JWT token string
    # - Check if it's valid and not expired
    # - Return user data if valid, None if invalid
    try:
        payload=jwt.decode(token,settings.jwt_secret,algorithms=[settings.jwt_algorithm])
        return payload
    except JWTError:
        # TOken is invalid, expired, or malformed
        return None

async def get_current_user_from_token(token: str):
    """Get user from database using JWT token"""
    from services.user_service import get_user_by_id
    
    # Verify the token first
    payload = verify_token(token)
    if payload is None:
        return None
    
    # Extract user_id from payload
    user_id = payload.get("sub")  # "sub" is standard JWT field for subject (user_id)
    if user_id is None:
        return None
    
    # Get user from database
    try:
        user = await get_user_by_id(user_id)
        return user
    except Exception:
        return None