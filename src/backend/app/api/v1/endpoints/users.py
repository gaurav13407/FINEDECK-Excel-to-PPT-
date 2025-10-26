# User Management Endpoints
# Handles user profile and account management
# - GET /users/me - Get current user profile
# - PUT /users/me - Update user profile information
# - DELETE /users/me - Delete user account
# - GET /users/stats - Get user usage statistics
# - PUT /users/password - Change user password
# - GET /users/conversion-history - Get user's conversion history
# - PUT /users/preferences - Update user preferences and settings


from fastapi import APIRouter, Depends, HTTPException, status,Query
from typing import Optional
import logging
import traceback

from services.user_service import(
    get_user_by_id,update_user_profile,deduct_user_credits,
    upgrade_subscription,get_user_usage_stats as service_get_user_usage_stats
)
from models.user import(
    UserResponse,UserUpdate,SubscriptionPlan,SubscriptionUpdate,
    PresentationUsage,SubscriptionResponse
)
from api.deps import get_current_active_user,get_db,get_pagination_params
router=APIRouter(tags=["User"])  # Remove prefix from here, it's added in api.py

@router.get("/me",response_model=UserResponse)
async def get_current_user_profile(
    current_user=Depends(get_current_active_user)
):
    """Get current user profile"""
    return UserResponse.model_validate(current_user)

@router.put("/me",response_model=UserResponse)
async def update_user_profile_endpoint(
    profile_updates:UserUpdate,
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    try:
        # Convert UserUpdate to dict, excluding None values
        profile_data = profile_updates.dict(exclude_unset=True)
        update_user=await update_user_profile(current_user.id, profile_data)
        if not update_user:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="User not found"
            )
        return UserResponse.model_validate(update_user)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error updating user profile: {str(e)}"
        )
    

@router.delete("/me")
async def delete_user_account(
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """Delete current user account"""
    try:
        deactivated_user=UserUpdate(is_active=False)
        deactivated_data = deactivated_user.dict(exclude_unset=True)
        update_user=await update_user_profile(current_user.id,deactivated_data)
        if not update_user:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="User not found"
            )
        return{
            "message":"User account deactivated successfully",
            "user_id" : str(current_user.id)
             }
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error deleting user account: {str(e)}"
        )
    

@router.get("/stats")
async def get_user_stats_endpoint(current_user=Depends(get_current_active_user),db=Depends(get_db)):
    """Get user usage statistics"""
    try:
        # Ensure we pass a string user_id to the service
        user_id_str = str(getattr(current_user, 'id', getattr(current_user, '_id', current_user)))
        stats = await service_get_user_usage_stats(user_id_str)

        # Defensive extraction of fields (models may differ between DB versions)
        subscription = getattr(current_user, 'subscription', None) or {}
        # subscription.plan may be an Enum or string
        plan = getattr(subscription, 'plan', None)
        if hasattr(plan, 'value'):
            plan = plan.value

        # presentation counts: support both presentation_created and presentations_created
        presentations_created = getattr(current_user, 'presentations_created', None)
        if presentations_created is None:
            presentations_created = getattr(current_user, 'presentation_created', 0)

        # presentation limit: try subscription.presentations_limit, then legacy field
        presentations_limit = getattr(subscription, 'presentations_limit', None) or getattr(current_user, 'presentations_limit', None)

        # credits from subscription (defensive)
        credits_used = getattr(subscription, 'monthly_credits_used', None)
        credits_limit = getattr(subscription, 'monthly_credits_limit', None)
        if credits_used is None or credits_limit is None:
            # fallback to stats returned by service
            credits_used = credits_used if credits_used is not None else stats.get('credits_used') if isinstance(stats, dict) else None
            credits_limit = credits_limit if credits_limit is not None else stats.get('credits_limit') if isinstance(stats, dict) else None

        credits_remaining = None
        if (credits_limit is not None) and (credits_used is not None):
            try:
                credits_remaining = int(credits_limit) - int(credits_used)
            except Exception:
                credits_remaining = None

        return {
            "user_id": str(current_user.id),
            "subscription_plan": plan,
            "presentations_created": presentations_created,
            "presentations_limit": presentations_limit,
            "credits_used": credits_used,
            "credits_limit": credits_limit,
            "credits_remaining": credits_remaining,
            "account_created_at": getattr(current_user, 'created_at', None),
            "last_login": getattr(current_user, 'last_login', None)
        }
    except Exception as e:
        # Log full traceback for debugging
        logging.exception("Error in get_user_usage_stats endpoint")
        traceback.print_exc()
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error retrieving user stats: {str(e)}"
        )


@router.get('/debug/current_user')
async def debug_current_user(current_user=Depends(get_current_active_user)):
    """Dev-only: return serialized current_user for debugging field names and values."""
    try:
        # Prefer pydantic v2 model_dump, fallback to dict/vars
        if hasattr(current_user, 'model_dump'):
            data = current_user.model_dump()
        else:
            try:
                data = dict(current_user)
            except Exception:
                # last resort
                data = {k: str(v) for k, v in getattr(current_user, '__dict__', {}).items()}
    except Exception as e:
        logging.exception('Failed to serialize current_user')
        data = {'error': str(e)}

    return { 'current_user': data }
    

@router.post("/subscription/upgrade")
async def upgrade_subscription(
    subscription_data:SubscriptionUpdate,
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """Upgrade user subscription plan"""
    current_plan_level={
        SubscriptionPlan.FREE:0,
        SubscriptionPlan.BASIC:1,
        SubscriptionPlan.PRO:2,
        SubscriptionPlan.ENTERPRISE:3
    }
    current_level=current_plan_level.get(current_user.subscription.plan,0)
    new_level=current_plan_level.get(subscription_data.new_plan,0)
    if new_level <= current_level:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="New subscription plan must be higher than current plan"
        )
    try:
        updated_user=await upgrade_subscription(current_user.id,subscription_data.new_plan)
        if not updated_user:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="User not found"
            )
        return {
            "message":f"Subscription upgraded to {subscription_data.new_plan} successfully",
            "previous_plan": current_user.subscription.plan.value,
            "new_plan": subscription_data.plan.value,
            "new_limits":{
                "presentations_limit":updated_user.presentations_limit,
                "monthly_credits_limit":updated_user.subscription.monthly_credits_limit
            }
        }
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error upgrading subscription: {str(e)}"
        )
    

@router.get("/subscription")
async def get_subscription_info(current_user=Depends(get_current_active_user)):
    """Get current user's subscription information"""
    subscription=current_user.subscription

    days_remaining=None
    # Support both `expires_at` (used by models) and legacy `expiry_at` if present
    expires = None
    try:
        expires = getattr(subscription, 'expires_at', None) or getattr(subscription, 'expiry_at', None)
    except Exception:
        expires = None

    if expires:
        # Normalize to aware datetime if possible; calculate remaining days defensively
        try:
            from datetime import datetime, timezone
            now_dt = datetime.now(timezone.utc)
            remaning = expires - now_dt
            days_remaining = max(0, remaning.days)
        except Exception:
            days_remaining = None

    # Defensive extraction for presentations_limit (may live on subscription or user)
    presentations_limit = getattr(subscription, 'presentations_limit', None) or getattr(current_user, 'presentations_limit', None)

    # If still None, try common alternatives or default to 0
    if presentations_limit is None:
        presentations_limit = getattr(subscription, 'presentationsLimit', None) or 0

    return SubscriptionResponse(
        plan=subscription.plan,
        status=subscription.status,
        price_per_month=getattr(subscription, 'price_per_month', 0.0),
        presentations_limit=presentations_limit,
        ai_features_enabled=getattr(subscription, 'ai_features_enabled', False),
        days_remaining=days_remaining
    )


@router.post("/credit/deduct")
async def deduct_credits(
    credits_amount:int=Query(...,gt=0,le=10,description="Number of credits to deduct"),db=Depends(get_db),current_user=Depends(get_current_active_user)
):
    """Deduct credits from user's account"""
    remaining_credits=(
        current_user.subscription.monthly_credits_limit - current_user.subscription.monthly_credits_used
    )
    if remaining_credits < credits_amount:
        raise HTTPException(
            status_code=status.HTTP_402_PAYMENT_REQUIRED,
            detail=f"Insufficient credits. Available: {remaining_credits}, Requested deduction: {credits_amount}"
        )
    
    try:
        updated_user=await deduct_user_credits(current_user.id,credits_amount)
        if not updated_user:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="User not found"
            )
        new_remaining=(
            updated_user.subscription.monthly_credits_limit - updated_user.subscription.monthly_credits_used
        )
        return {
            "message":f"Deducted {credits_amount} credits successfully",
            "credits_used": updated_user.subscription.monthly_credits_used,
            "credits_remaining": new_remaining
        }
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error deducting credits: {str(e)}"
        )
    

@router.get("/conversion-history")
async def get_conversion_history(
    pagination=Depends(get_pagination_params),
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """
    Get user's conversion history with pagination
    
    - **skip**: Number of records to skip (default: 0)
    - **limit**: Number of records to return (default: 20, max: 100)
    """
    try:
        # This would typically fetch from conversions collection
        # For now, return basic structure
        return {
            "conversions": [],
            "total_count": 0,
            "skip": pagination["skip"],
            "limit": pagination["limit"],
            "message": "Conversion history feature coming soon"
        }
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Failed to get conversion history: {str(e)}"
        )


# Backwards-compatible aliases: some frontends call /users/profile — keep compatibility
@router.get("/profile", response_model=UserResponse)
async def get_profile_alias(current_user=Depends(get_current_active_user)):
    """Alias for GET /users/me for backward compatibility with older frontends."""
    return UserResponse.model_validate(current_user)


@router.put("/profile", response_model=UserResponse)
async def update_profile_alias(
    profile_updates: UserUpdate,
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """Alias for PUT /users/me to support older frontend paths."""
    # Reuse existing update logic
    try:
        profile_data = profile_updates.dict(exclude_unset=True)
        update_user = await update_user_profile(current_user.id, profile_data)
        if not update_user:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="User not found"
            )
        return UserResponse.model_validate(update_user)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error updating user profile: {str(e)}"
        )