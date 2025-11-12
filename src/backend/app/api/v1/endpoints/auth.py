# Authentication Endpoints
# Handles all user authentication and authorization
# - POST /auth/signup - User registration with email verification
# - POST /auth/login - User login with JWT token generation
# - POST /auth/refresh - JWT token refresh functionality
# - POST /auth/logout - User logout with token blacklisting
# - POST /auth/forgot-password - Password reset request
# - POST /auth/reset-password - Password reset confirmation
# - GET /auth/me - Get current user profile information

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.security import OAuth2PasswordRequestForm
from datetime import timedelta

from core.security import create_access_token, verify_password
from core.config import settings
from services.user_service import create_user, get_user_by_email,authenticate_user
from models.user import UserCreate, UserResponse,UserLogin,Token
from api.deps import get_current_active_user,get_db
from pydantic import BaseModel
from datetime import datetime
from core.database import get_collection
from core.security import get_password_hash

router = APIRouter(tags=["Authentication"])
@router.post("/signup", response_model=UserResponse,status_code=status.HTTP_201_CREATED)
async def signup(user_data: UserCreate,db=Depends(get_db)):
    """ Register New User ACcount"""
    existing_user=await get_user_by_email(user_data.email)
    if existing_user:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Email already registered"
        )
    try:
        new_user=await create_user(user_data)
        return UserResponse.model_validate(new_user)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=str(e)
        )
    

@router.post("/login", response_model=Token)
async def login(form_data: OAuth2PasswordRequestForm = Depends()):
    """Login with email and password to get JWT token"""
    
    # Debug logging
    print(f"🔍 Login attempt:")
    print(f"   - Email/Username: {form_data.username}")
    print(f"   - Password provided: {'Yes' if form_data.password else 'No'}")
    print(f"   - Password length: {len(form_data.password) if form_data.password else 0}")
    
    user=await authenticate_user(form_data.username,form_data.password)
    print(f"   - Authentication result: {'Success' if user else 'Failed'}")
    
    if not user:
        print(f"❌ Authentication failed for {form_data.username}")
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect email or password",
            headers={"WWW-Authenticate": "Bearer"},
        )
    if not user.is_active:
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="Inactive user"
        )
    
    # Use configured JWT expiration time (7 days by default)
    access_token_expires=timedelta(minutes=settings.jwt_expiration_minutes)
    access_token=create_access_token(
        data={"sub": user.email, "user_id": str(user.id)},
        expires_delta=access_token_expires
    )
    return Token(
        access_token=access_token,
        token_type="bearer",
        user=UserResponse.model_validate(user)
    )


@router.post("/refresh",response_model=Token)
async def refresh_token(current_user=Depends(get_current_active_user)):
    """Refresh JWT token for authenticated user"""
    # Use configured JWT expiration time (7 days by default)
    access_token_expires=timedelta(minutes=settings.jwt_expiration_minutes)
    access_token=create_access_token(
        data={"sub": current_user.email, "user_id": str(current_user.id)},
        expires_delta=access_token_expires
    )
    return Token(
        access_token=access_token,
        token_type="bearer",
        user=UserResponse.model_validate(current_user)
    )

@router.get("/me", response_model=UserResponse)
async def get_current_user_profile(
    current_user=Depends(get_current_active_user)
):
    """Get current authenticated user's profile"""
    return UserResponse.model_validate(current_user)


@router.post("/logout")
async def logout():
    """Logout user (token blacklisting can be implemented here)"""
    return {"message": "Logout successful"}

@router.post("/change-password")
async def change_password(
    current_password: str,
    new_password: str,
    current_user=Depends(get_current_active_user),
):
    """Change password for authenticated user"""
    if not verify_password(current_password, current_user.password_hash):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Current password is incorrect"
        )
    try:
        # Update password hash in database
        users = get_collection("users")
        new_hash = get_password_hash(new_password)
        await users.update_one({"_id": current_user.id}, {"$set": {"password_hash": new_hash}})
        return {"message": "Password changed successfully"}
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to update password")



class ResetPasswordRequest(BaseModel):
    email: str
    code: str
    new_password: str
    confirm_password: str


@router.post("/reset-password")
async def reset_password(payload: ResetPasswordRequest):
    """Reset password using a one-time code sent to the user's email.

    Steps:
    - Validate passwords match
    - Find activation code with purpose 'password_reset'
    - Verify not expired and not redeemed
    - Ensure new password is not the same as current password
    - Update user's password_hash and mark code redeemed
    """
    if payload.new_password != payload.confirm_password:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Passwords do not match")

    codes = get_collection("activation_codes")
    users = get_collection("users")
    now = datetime.utcnow()

    q = {
        "code": payload.code,
        "purpose": "password_reset",
        "target_email": payload.email,
        "redeemed": False,
        "expires_at": {"$gt": now}
    }
    doc = await codes.find_one(q)
    if not doc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="Invalid or expired code")

    user = await users.find_one({"email": payload.email})
    if not user:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="User not found")

    # Prevent reusing the same password
    if user.get("password_hash") and verify_password(payload.new_password, user.get("password_hash")):
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail="New password cannot be the same as the current password")

    # All good: update password and mark code redeemed
    new_hash = get_password_hash(payload.new_password)
    await users.update_one({"_id": user.get("_id")}, {"$set": {"password_hash": new_hash}})
    await codes.update_one({"_id": doc["_id"]}, {"$set": {"redeemed": True, "redeemed_at": now, "redeemed_by": payload.email}})

    return {"ok": True, "message": "Password has been reset"}