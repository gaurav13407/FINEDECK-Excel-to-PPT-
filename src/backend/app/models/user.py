# User MongoDB Model
# Defines the user data structure and validation
# - User profile information (name, email, password hash)
# - Account status and verification state
# - Subscription plan and usage limits
# - User preferences and settings
# - Timestamps for account creation and last login
# - File upload history and storage usage
# - Pydantic models for request/response validation
from pydantic import BaseModel,EmailStr,Field
from typing import Optional
from datetime import datetime
from bson import ObjectId
from enum import Enum

## MongoDB uses ObjectedID,but APIs use strings
class PyObjectId(ObjectId):
    @classmethod
    def __get_pydantic_core_schema__(cls, source_type, handler):
        from pydantic_core import core_schema
        return core_schema.no_info_plain_validator_function(cls.validate)

    @classmethod
    def validate(cls, v):
        if isinstance(v, ObjectId):
            return v
        if isinstance(v, str) and ObjectId.is_valid(v):
            return ObjectId(v)
        raise ValueError("Invalid ObjectId")
    
    @classmethod
    def __get_validators__(cls):
        yield cls.validate
    
class UserCreate(BaseModel):
    name:str=Field(...,min_length=2,max_length=100,description="User's full name")
    email:EmailStr=Field(...,description="User's email address")
    password:str=Field(...,min_length=6,max_length=100,description="User's password")


class UserLogin(BaseModel):
    email:EmailStr=Field(...,description="User's email address")
    password:str=Field(...,min_length=1,description="User's password")

class UserResponse(BaseModel):
    id:PyObjectId=Field(default_factory=PyObjectId,alias="_id")
    name:str
    email:EmailStr
    subscription:"SubscriptionDetails"  # Forward reference with quotes
    is_active:bool=True
    created_at:datetime
    last_login:Optional[datetime]=None
    presentations_created:int=0
    # presentations_limit:int=7
    
    model_config = {
        "from_attributes": True,
        "arbitrary_types_allowed": True,
        "populate_by_name": True,
        "json_encoders": {ObjectId: str}
    }

class UserInDB(UserResponse):
    password_hash:str=Field(...,description="Hashed password")

class UserUpdate(BaseModel):
    """Model for updating user profile information"""
    name: Optional[str] = Field(None, min_length=2, max_length=100, description="Updated full name")
    email: Optional[EmailStr] = Field(None, description="Updated email address")
    is_active: Optional[bool] = Field(None, description="Account status")
    
    model_config = {
        "from_attributes": True,
        "populate_by_name": True
    }

class Token(BaseModel):
    access_token:str
    token_type:str="bearer"
    user:UserResponse

class SubscriptionPlan(str,Enum):
    FREE="free"
    BASIC="basic"
    PRO="pro"
    AI_PRO="ai_pro"  # Updated from ENTERPRISE to match our tier naming
    ENTERPRISE="enterprise"  # Keep for backward compatibility

class SubscriptionStatus(str,Enum):
    ACTIVE="active"
    CANCELLED="cancelled"
    PAST_DUE="past_due"
    TRIAL="trial"


class SubscriptionUpdate(BaseModel):
    plan:SubscriptionPlan
    auto_renew:Optional[bool]=None


class SubscriptionResponse(BaseModel):
    plan:SubscriptionPlan
    status:SubscriptionStatus
    price_per_month:float
    presentations_limit:int
    ai_features_enabled:bool
    days_remaining:Optional[int]=None


class AIRequest(BaseModel):
    prompt:str=Field(...,min_length=1,max_length=500,description="AI processing request")
    excel_data_summary:Optional[str]=None
    template_preference:Optional[str]=None

class AIResponse(BaseModel):
    suggestions:list[str]
    insights:list[str]
    recommended_charts:list[str]
    processing_time:float


class PresentationUsage(BaseModel):
    user_id:PyObjectId
    created_at:datetime
    file_size_mb:float
    processing_time:float
    ai_features_used:bool=False

class UsageStats(BaseModel):
    current_month_usage:int
    remaining_presentations:int
    usage_percentage:float
    ai_requests_used:int=0

## Business logic
def can_create_presentation(user:UserInDB)->bool:
    config=PLAN_CONFIGS[user.subscription.plan]
    return user.presentations_created < config["presentations_limit"]


def has_ai_access(user:UserInDB)->bool:
    return user.subscription.plan == SubscriptionPlan.ENTERPRISE and user.subscription.ai_features_enabled


# Calculate monthly Revenu
def calculate_mrr(users:list[UserInDB])->float:
    return sum(PLAN_CONFIGS[user.subscription.plan]["price"]for user in users if user.subscription.status == SubscriptionStatus.ACTIVE)

class TemplateCategory(str,Enum):
    BASIC="basic"
    PROFESSIONAL="professional"
    PREMIUM="premium"
    CUSTOM="custom"


def has_template_access(user:UserInDB,template_category:TemplateCategory)->bool:
    config=PLAN_CONFIGS[user.subscription.plan]
    template_access=config["template_access"]

    if template_category == TemplateCategory.BASIC:
        return template_access["basic_templates"]
    elif template_category == TemplateCategory.PROFESSIONAL:
        return template_access["professional_templates"]
    elif template_category == TemplateCategory.PREMIUM:
        return template_access["premium_templates"]
    elif template_category == TemplateCategory.CUSTOM:
        return template_access["custom_templates"]
    
    return False


def can_upload_template(user:UserInDB)->bool:
    config=PLAN_CONFIGS[user.subscription.plan]
    return config["template_access"]["template_upload"]


def get_available_templates(user:UserInDB)->list[TemplateCategory]:
    available=[]
    config=PLAN_CONFIGS[user.subscription.plan]
    template_access=config["template_access"]

    if template_access["basic_templates"]:
        available.append(TemplateCategory.BASIC)
    if template_access["professional_templates"]:
        available.append(TemplateCategory.PROFESSIONAL)
    if template_access["premium_templates"]:
        available.append(TemplateCategory.PREMIUM)
    if template_access["custom_templates"]:
        available.append(TemplateCategory.CUSTOM)

    return available

class TemplateAccess(BaseModel):
    basic_templates:bool=True # ALL Plans get basic
    professional_templates:bool=False # pro+Enterprise
    premium_templates:bool=False #Enterprise
    custom_template_upload:bool=False #Enterprise
    template_upload:bool=False # Enterprise only



class SubscriptionDetails(BaseModel):
    plan:SubscriptionPlan=SubscriptionPlan.BASIC
    status:SubscriptionStatus=SubscriptionStatus.TRIAL
    price_per_month:float=Field(...,description="Monthly price in USD")
    presentations_limit:int=Field(...,description="Monthly presentation limit")

    #New: Credit System
    monthly_credits_limit:int=Field(...,description="Monthly credit limit")
    monthly_credits_used:int=Field(0,description="Credits used this month")
    credits_reset_date:Optional[datetime]=Field(...,description="Date when credits reset")

    ai_features_enabled:bool=False
    started_at:datetime
    expires_at:Optional[datetime]=None
    auto_renew:bool=True

## Add to UsageSats model
class UsageStats(BaseModel):
    current_month_usage:int
    remaining_presentations:int
    usage_percentage:float
    ai_requests_used:int=0

    # New credit fields
    monthly_credits_used:int=0
    monthly_credits_limit:int
    remaining_credits:int=0
    credits_usage_percentage:float=0.0



# New: Credits Businedd logic Functions
def has_credits_remaning(user:UserInDB,credits_needed:int=1)->bool:
    """Checks if user has enough credits to create a presentation"""
    config=PLAN_CONFIGS[user.subscription.plan]
    credit_limit = config["monthly_credits_limit"]
    
    # Check if unlimited credits (-1 or 200+)
    if credit_limit == -1 or credit_limit >= 200:
        return True  # Unlimited credits
    
    total_used=user.subscription.monthly_credits_used+credits_needed
    return total_used <= credit_limit


def deduct_credits(user:UserInDB,credits_used:int)-> bool:
    """Deduct credits from user's subscription"""
    if has_credits_remaning(user,credits_used):
        user.subscription.monthly_credits_used += credits_used
        return True
    return False


def get_credit_usage_stats(user:UserInDB)->dict:
    """Get user's credit usage stats"""
    config=PLAN_CONFIGS[user.subscription.plan]
    limit=config["monthly_credits_limit"]
    used=user.subscription.monthly_credits_used
    return {
        "credits_used":used,
        "credits_limit":limit,
        "credits_remaining":limit - used,
        "usage_percentage":(used / limit * 100) if limit > 0 else 0
    }



PLAN_CONFIGS={
    SubscriptionPlan.FREE: {
        "price": 0.00,
        "presentations_limit": 1,
        "ai_features": [],
        "ai_titles": False,
        "ai_template_selection": False,
        "ai_summaries": False,
        "ai_insights": False,
        "ai_layout": False,
        "ai_chart_recommendations": False,
        "monthly_credits_limit": 1,  # Aligned with PPT limit (1 credit = 1 PPT)
        "max_sheets": 1,
        "template_access": {
            "basic_templates": True,
            "professional_templates": False,
            "premium_templates": False,
            "custom_templates": False,
            "custom_template_upload": False,
            "template_upload": False
        },
        "name": "Free Plan",
        "features": ["1 presentation/month", "Basic template only", "Single sheet support", "No AI features"],
    },
    SubscriptionPlan.BASIC:{
        "price":25.00,
        "presentations_limit":7,
        "ai_features":["title"],
        "ai_titles": True,
        "ai_template_selection": False,
        "ai_summaries": False,
        "ai_insights": False,
        "ai_layout": False,
        "ai_chart_recommendations": False,
        "monthly_credits_limit":7,  # FIXED: Aligned with PPT limit (7 credits = 7 PPTs)
        "max_sheets": 5,
        "template_access":{
            "basic_templates":True,
            "professional_templates":False,
            "premium_templates":False,
            "custom_templates":False,
            "custom_template_upload":False,
            "template_upload":False
        },
        "name":"Basic Plan",
        "features":["7 presentations/month", "Basic template", "AI-generated titles", "Multi-sheet support (5 max)", "Standard processing"],
    },
    SubscriptionPlan.PRO:{
        "price":49.00,
        "presentations_limit":15,
        "ai_features":["title", "template_selection"],
        "ai_titles": True,
        "ai_template_selection": True,
        "ai_summaries": False,
        "ai_insights": False,
        "ai_layout": False,
        "ai_chart_recommendations": False,
        "monthly_credits_limit":15,  # FIXED: Aligned with PPT limit (15 credits = 15 PPTs)
        "max_sheets": 20,
        "template_access":{
            "basic_templates":True,
            "professional_templates":True,
            "premium_templates":False,
            "custom_templates":False,
            "custom_template_upload":False,
            "template_upload":False
        },
        "name":"Pro Plan",
        "features":["15 presentations/month", "All 10 professional templates", "AI titles + template selection", "Multi-sheet support (20 max)", "Faster processing"],
    },
    SubscriptionPlan.AI_PRO:{
        "price":99.00,
        "presentations_limit":-1,  # Unlimited
        "ai_features":["title", "summary", "insights", "template_selection", "layout", "chart_type"],
        "ai_titles": True,
        "ai_template_selection": True,
        "ai_summaries": True,
        "ai_insights": True,
        "ai_layout": True,
        "ai_chart_recommendations": True,
        "monthly_credits_limit":-1,  # FIXED: Unlimited (was 1000, now -1 to match unlimited PPTs)
        "max_sheets": -1,  # Unlimited
        "template_access":{
            "basic_templates":True,
            "professional_templates":True,
            "premium_templates":True,
            "custom_templates":True,
            "custom_template_upload":True,
            "template_upload":True
        },
        "name":"AI Pro Plan",
        "features":["Unlimited presentations", "All 10 professional templates", "All 6 AI features", "Unlimited sheets", "AI-generated titles", "AI summaries", "AI insights (5 bullets)", "AI layout optimization", "AI chart recommendations", "Priority support"]
    },
    SubscriptionPlan.ENTERPRISE:{
        "price":99.99,
        "presentations_limit":1000,
        "ai_features":["title", "summary", "insights", "template_selection", "layout", "chart_type"],
        "ai_titles": True,
        "ai_template_selection": True,
        "ai_summaries": True,
        "ai_insights": True,
        "ai_layout": True,
        "ai_chart_recommendations": True,
        "monthly_credits_limit":1000,  # Aligned with PPT limit (1000 credits = 1000 PPTs)
        "max_sheets": -1,
        "template_access":{
            "basic_templates":True,
            "professional_templates":True,
            "premium_templates":True,
            "custom_templates":True,
            "custom_template_upload":True,
            "template_upload":True
        },
        "name":"Enterprise Plan (Legacy)",
        "features":["1000 presentations/month", "All templates", "All AI features", "Priority support"]
    },
}

# Rebuild models to resolve forward references
UserResponse.model_rebuild()
UserInDB.model_rebuild()