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

from services.user_service import(
    get_user_by_id,update_user_profile,deduct_user_credits,
    upgrade_subscription,get_user_usage_stats
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
async def get_user_usage_stats(current_user=Depends(get_current_active_user),db=Depends(get_db)):
    """Get user usage statistics"""
    try:
        stats=await get_user_usage_stats(current_user.id)
        return {
            "user_id": str(current_user.id),
            "subscription_plan": current_user.subscription.plan,
            "presentation_created":current_user.presentation_created,
            "presentation_limit": current_user.presentations_limit,
            "credits_used": current_user.subscription.monthly_credits_used,
            "credits_limit": current_user.subscription.monthly_credits_limit,
            "credits_remaining":(
                current_user.subscription.monthly_credits_limit - current_user.subscription.monthly_credits_used
            ),
            "account_created_at": current_user.created_at,
            "last_login": current_user.last_login
        }
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error retrieving user stats: {str(e)}"
        )
    

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
    if subscription.expiry_at:
        from datetime import datetime,timezone,UTC
        remaning=subscription.expiry_at - datetime.now(UTC)
        days_remaining=max(0,remaning.days)

    return SubscriptionResponse(
        plan=subscription.plan,
        status=subscription.status,
        price_per_month=subscription.price_per_month,
        presentations_limit=current_user.presentations_limit,
        ai_features_enabled=subscription.ai_features_enabled,
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