"""
Plan Upgrade Code System
========================
Allows admin to send upgrade codes to customers via email.
Customers can redeem codes to upgrade their subscription plan.

Flow:
1. Admin generates upgrade code for specific plan (BASIC, PRO, AI_PRO)
2. System sends email to customer with upgrade code
3. Customer enters code in their account settings
4. System validates code and upgrades their plan
"""

from fastapi import APIRouter, Depends, HTTPException, status, Request
from pydantic import BaseModel, EmailStr
from datetime import datetime, timedelta
from typing import Optional, List
import secrets
import string
from bson import ObjectId

from core.database import get_collection
from core.config import settings
from services.emails_utils import send_email
from api.deps import get_current_active_user
from models.user import UserInDB, SubscriptionPlan

router = APIRouter()

# Upgrade code configuration
UPGRADE_CODE_LENGTH = 12
UPGRADE_CODE_EXPIRY_DAYS = 30  # Codes expire after 30 days


def generate_upgrade_code() -> str:
    """Generate a unique upgrade code (e.g., BASIC-A1B2C3D4E5F6)"""
    characters = string.ascii_uppercase + string.digits
    code_suffix = ''.join(secrets.choice(characters) for _ in range(UPGRADE_CODE_LENGTH))
    return code_suffix


def create_upgrade_email_html(customer_email: str, plan: str, code: str, expires_at: datetime) -> str:
    """Create HTML email for upgrade code"""
    
    plan_names = {
        "BASIC": "Basic Plan",
        "PRO": "Pro Plan", 
        "AI_PRO": "AI Pro Plan"
    }
    
    plan_features = {
        "BASIC": [
            "✓ 15 presentations per month",
            "✓ 5 template designs",
            "✓ AI-powered slide titles",
            "✓ Basic charts and graphs"
        ],
        "PRO": [
            "✓ 50 presentations per month",
            "✓ 10 premium templates",
            "✓ AI titles & summaries",
            "✓ Advanced charts with legends",
            "✓ Custom branding"
        ],
        "AI_PRO": [
            "✓ 200 presentations per month",
            "✓ All premium templates",
            "✓ Full AI suite (insights, predictions, recommendations)",
            "✓ Smart chart analyzer",
            "✓ Priority support",
            "✓ API access"
        ]
    }
    
    plan_name = plan_names.get(plan, plan)
    features = plan_features.get(plan, [])
    
    html = f"""
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 600px;
            margin: 0 auto;
            padding: 20px;
        }}
        .container {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 10px;
            padding: 30px;
            color: white;
        }}
        .logo {{
            font-size: 32px;
            font-weight: bold;
            margin-bottom: 20px;
        }}
        .content {{
            background: white;
            color: #333;
            border-radius: 8px;
            padding: 30px;
            margin-top: 20px;
        }}
        .code-box {{
            background: #f8f9fa;
            border: 2px dashed #667eea;
            border-radius: 8px;
            padding: 20px;
            text-align: center;
            margin: 30px 0;
        }}
        .code {{
            font-size: 28px;
            font-weight: bold;
            color: #667eea;
            letter-spacing: 2px;
            font-family: 'Courier New', monospace;
        }}
        .features {{
            background: #f8f9fa;
            border-radius: 8px;
            padding: 20px;
            margin: 20px 0;
        }}
        .feature {{
            padding: 8px 0;
            font-size: 16px;
        }}
        .cta-button {{
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px 40px;
            border-radius: 5px;
            text-decoration: none;
            font-weight: bold;
            margin: 20px 0;
        }}
        .footer {{
            text-align: center;
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #eee;
            color: #666;
            font-size: 14px;
        }}
        .expiry {{
            color: #dc3545;
            font-weight: bold;
            margin-top: 15px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">🎯 FinDeck</div>
        <h1 style="margin: 0;">You've Been Upgraded!</h1>
        <p style="font-size: 18px; margin-top: 10px;">Welcome to {plan_name}</p>
    </div>
    
    <div class="content">
        <h2>🎉 Congratulations!</h2>
        <p>Your FinDeck account has been upgraded to <strong>{plan_name}</strong>!</p>
        
        <div class="code-box">
            <div style="font-size: 14px; color: #666; margin-bottom: 10px;">Your Upgrade Code:</div>
            <div class="code">{code}</div>
            <div class="expiry">⏰ Expires: {expires_at.strftime('%B %d, %Y')}</div>
        </div>
        
        <h3>📦 What's Included:</h3>
        <div class="features">
            {''.join(f'<div class="feature">{feature}</div>' for feature in features)}
        </div>
        
        <h3>🚀 How to Activate:</h3>
        <ol>
            <li>Log in to your FinDeck account</li>
            <li>Go to <strong>Account Settings</strong> → <strong>Subscription</strong></li>
            <li>Click <strong>"Redeem Upgrade Code"</strong></li>
            <li>Enter the code above</li>
            <li>Start creating amazing presentations!</li>
        </ol>
        
        <div style="text-align: center;">
            <a href="{settings.FRONTEND_URL}/dashboard" class="cta-button">
                Activate Now →
            </a>
        </div>
        
        <div class="footer">
            <p>This code is valid for 30 days and can only be used once.</p>
            <p>Need help? Contact us at support@findeck.com</p>
            <p style="margin-top: 20px; color: #999;">
                © 2025 FinDeck. All rights reserved.
            </p>
        </div>
    </div>
</body>
</html>
"""
    return html


class GenerateUpgradeCodeRequest(BaseModel):
    email: EmailStr
    plan: str  # "BASIC", "PRO", or "AI_PRO"
    duration_months: Optional[int] = None  # Optional: specific duration
    notes: Optional[str] = None  # Internal notes about why code was generated


class RedeemUpgradeCodeRequest(BaseModel):
    code: str


class UpgradeCodeResponse(BaseModel):
    code: str
    plan: str
    customer_email: str
    generated_at: datetime
    expires_at: datetime
    is_redeemed: bool
    redeemed_at: Optional[datetime] = None
    redeemed_by: Optional[str] = None


@router.post("/generate-upgrade-code", status_code=status.HTTP_201_CREATED)
async def generate_upgrade_code_for_customer(
    request: GenerateUpgradeCodeRequest,
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Admin endpoint: Generate upgrade code and send via email
    
    This creates a unique code that allows the customer to upgrade their plan.
    The code is sent to their email and can be redeemed once.
    """
    
    # Validate plan
    valid_plans = ["BASIC", "PRO", "AI_PRO"]
    if request.plan.upper() not in valid_plans:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Invalid plan. Must be one of: {', '.join(valid_plans)}"
        )
    
    plan = request.plan.upper()
    
    # Generate unique code
    code = generate_upgrade_code()
    full_code = f"{plan}-{code}"
    
    # Calculate expiry
    now = datetime.utcnow()
    expires_at = now + timedelta(days=UPGRADE_CODE_EXPIRY_DAYS)
    
    # Save to database
    codes_collection = get_collection("upgrade_codes")
    
    code_document = {
        "code": full_code,
        "plan": plan,
        "customer_email": request.email,
        "generated_by": str(current_user.id) if hasattr(current_user, 'id') else str(current_user._id),
        "generated_by_email": current_user.email,
        "generated_at": now,
        "expires_at": expires_at,
        "is_redeemed": False,
        "redeemed_at": None,
        "redeemed_by": None,
        "duration_months": request.duration_months,
        "notes": request.notes,
        "meta": {
            "ip": None,  # Can be captured from request
            "user_agent": None
        }
    }
    
    result = await codes_collection.insert_one(code_document)
    
    # Send email to customer
    try:
        email_html = create_upgrade_email_html(
            customer_email=request.email,
            plan=plan,
            code=full_code,
            expires_at=expires_at
        )
        
        email_sent = await send_email(
            to_email=request.email,
            subject=f"🎉 Your FinDeck {plan.replace('_', ' ').title()} Upgrade Code",
            html_content=email_html
        )
        
        if not email_sent:
            # Code was created but email failed - still return success
            # Admin can manually send the code
            print(f"⚠️ Warning: Email failed to send to {request.email}")
            
    except Exception as e:
        print(f"❌ Error sending email: {str(e)}")
        # Continue anyway - code is still valid
    
    return {
        "success": True,
        "message": f"Upgrade code generated and sent to {request.email}",
        "code": full_code,
        "plan": plan,
        "expires_at": expires_at.isoformat(),
        "customer_email": request.email,
        "code_id": str(result.inserted_id)
    }


@router.post("/redeem-upgrade-code")
async def redeem_upgrade_code(
    request: RedeemUpgradeCodeRequest,
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Customer endpoint: Redeem upgrade code to upgrade plan
    
    User enters the code they received via email to upgrade their account.
    Code can only be used once and must not be expired.
    """
    
    codes_collection = get_collection("upgrade_codes")
    users_collection = get_collection("users")
    
    # Find the code
    code_doc = await codes_collection.find_one({"code": request.code.upper()})
    
    if not code_doc:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail="Invalid upgrade code. Please check and try again."
        )
    
    # Check if already redeemed
    if code_doc.get("is_redeemed"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="This code has already been used."
        )
    
    # Check if expired
    if datetime.utcnow() > code_doc["expires_at"]:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="This code has expired. Please contact support for a new code."
        )
    
    # Verify email matches (optional - can be removed if codes are transferable)
    if code_doc["customer_email"].lower() != current_user.email.lower():
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail="This code was issued to a different email address."
        )
    
    # Get the plan from code
    plan_name = code_doc["plan"]
    
    # Map to SubscriptionPlan enum
    plan_mapping = {
        "BASIC": SubscriptionPlan.BASIC,
        "PRO": SubscriptionPlan.PRO,
        "AI_PRO": SubscriptionPlan.AI_PRO
    }
    
    new_plan = plan_mapping.get(plan_name)
    
    if not new_plan:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Invalid plan type in code."
        )
    
    # Calculate new expiry date
    now = datetime.utcnow()
    duration_months = code_doc.get("duration_months", 1)  # Default 1 month
    
    # For permanent upgrades, set far future date
    if duration_months is None or duration_months == 0:
        subscription_ends_at = now + timedelta(days=365 * 10)  # 10 years
    else:
        subscription_ends_at = now + timedelta(days=30 * duration_months)
    
    # Update user's subscription
    user_obj_id = ObjectId(str(current_user.id)) if hasattr(current_user, 'id') else ObjectId(str(current_user._id))
    
    update_result = await users_collection.update_one(
        {"_id": user_obj_id},
        {
            "$set": {
                "subscription.plan": new_plan,
                "subscription.status": "active",
                "subscription.starts_at": now,
                "subscription.ends_at": subscription_ends_at,
                "subscription.upgraded_at": now,
                "subscription.upgrade_method": "code",
                "subscription.upgrade_code": request.code.upper(),
                "updated_at": now
            }
        }
    )
    
    if update_result.modified_count == 0:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Failed to update subscription. Please contact support."
        )
    
    # Mark code as redeemed
    await codes_collection.update_one(
        {"code": request.code.upper()},
        {
            "$set": {
                "is_redeemed": True,
                "redeemed_at": now,
                "redeemed_by": str(user_obj_id),
                "redeemed_by_email": current_user.email
            }
        }
    )
    
    return {
        "success": True,
        "message": f"Congratulations! Your account has been upgraded to {plan_name.replace('_', ' ').title()}",
        "new_plan": plan_name,
        "subscription_ends_at": subscription_ends_at.isoformat(),
        "features_unlocked": get_plan_features(new_plan)
    }


@router.get("/my-upgrade-codes")
async def get_my_upgrade_codes(
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Get all upgrade codes sent to current user's email
    Shows both redeemed and unredeemed codes
    """
    
    codes_collection = get_collection("upgrade_codes")
    
    codes = await codes_collection.find(
        {"customer_email": current_user.email}
    ).sort("generated_at", -1).to_list(length=100)
    
    return {
        "codes": [
            {
                "code": code["code"],
                "plan": code["plan"],
                "generated_at": code["generated_at"].isoformat(),
                "expires_at": code["expires_at"].isoformat(),
                "is_redeemed": code.get("is_redeemed", False),
                "redeemed_at": code.get("redeemed_at").isoformat() if code.get("redeemed_at") else None,
                "is_expired": datetime.utcnow() > code["expires_at"],
                "can_redeem": not code.get("is_redeemed", False) and datetime.utcnow() <= code["expires_at"]
            }
            for code in codes
        ]
    }


@router.get("/admin/all-upgrade-codes")
async def get_all_upgrade_codes(
    current_user: UserInDB = Depends(get_current_active_user),
    include_redeemed: bool = True,
    include_expired: bool = False
):
    """
    Admin endpoint: Get all upgrade codes
    Useful for tracking and management
    """
    
    # TODO: Add admin role check
    # For now, any authenticated user can access (change in production)
    
    codes_collection = get_collection("upgrade_codes")
    
    # Build filter
    filter_query = {}
    
    if not include_redeemed:
        filter_query["is_redeemed"] = False
    
    if not include_expired:
        filter_query["expires_at"] = {"$gte": datetime.utcnow()}
    
    codes = await codes_collection.find(filter_query).sort("generated_at", -1).to_list(length=1000)
    
    return {
        "total_codes": len(codes),
        "codes": [
            {
                "code": code["code"],
                "plan": code["plan"],
                "customer_email": code["customer_email"],
                "generated_by": code.get("generated_by_email"),
                "generated_at": code["generated_at"].isoformat(),
                "expires_at": code["expires_at"].isoformat(),
                "is_redeemed": code.get("is_redeemed", False),
                "redeemed_at": code.get("redeemed_at").isoformat() if code.get("redeemed_at") else None,
                "redeemed_by_email": code.get("redeemed_by_email"),
                "notes": code.get("notes")
            }
            for code in codes
        ]
    }


def get_plan_features(plan: SubscriptionPlan) -> List[str]:
    """Get list of features for a plan"""
    features = {
        SubscriptionPlan.BASIC: [
            "15 presentations per month",
            "5 template designs",
            "AI-powered slide titles",
            "Basic charts and graphs"
        ],
        SubscriptionPlan.PRO: [
            "50 presentations per month",
            "10 premium templates",
            "AI titles & summaries",
            "Advanced charts with legends",
            "Custom branding"
        ],
        SubscriptionPlan.AI_PRO: [
            "200 presentations per month",
            "All premium templates",
            "Full AI suite (insights, predictions, recommendations)",
            "Smart chart analyzer",
            "Priority support",
            "API access"
        ]
    }
    
    return features.get(plan, [])
