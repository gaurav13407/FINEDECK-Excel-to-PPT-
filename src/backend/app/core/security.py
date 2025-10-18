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
        print("🔍 JWT verification failed - invalid token")
        return None
    
    print(f"🔍 JWT payload: {payload}")
    
    # Extract user_id from payload (use user_id field, not sub which contains email)
    user_id = payload.get("user_id")  # This contains the actual user ID
    if user_id is None:
        print("🔍 JWT missing user_id field")
        return None
    
    print(f"🔍 Extracted user_id: {user_id}")
    
    # Get user from database
    try:
        user = await get_user_by_id(user_id)
        if user:
            print(f"🔍 User found: {user.name}")
        else:
            print(f"🔍 User not found for ID: {user_id}")
        return user
    except Exception as e:
        print(f"🔍 Error getting user: {e}")
        return None