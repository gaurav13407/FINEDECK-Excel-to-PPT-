# User Service Layer
# Business logic for user management
# - User registration and profile creation
# - Password hashing and verification
# - Email verification and password reset
# - User authentication and session management
# - Profile updates and account settings
# - Usage tracking and subscription management
# - Account deletion and data cleanup


from typing import Optional,Dict,List,Any
from datetime import datetime,timedelta
from bson import ObjectId
from motor.motor_asyncio import AsyncIOMotorDatabase

# Import from models
from models.user import UserCreate,UserLogin,UserInDB,UserResponse,SubscriptionPlan,SubscriptionStatus,TemplateCategory
from core.security import get_password_hash,verify_password
from core.database import get_collection,get_database
from models.user import has_credits_remaning,deduct_credits,get_credit_usage_stats

async def create_user(user_data:UserCreate)->UserInDB:
    """Create a new user account with default subscription"""
   
    existing_user=await get_user_by_email(user_data.email)
    if existing_user:
        raise Exception(f"User with email {user_data.email} already exists.")
    
    password_hash=get_password_hash(user_data.password)

    user_doc={
        "name": user_data.name,
        "email": user_data.email,
        "password_hash": password_hash,
        "is_active": True,
        "subscription":{
            "plan": SubscriptionPlan.FREE,
            "status": SubscriptionStatus.ACTIVE,
            "price_per_month": 0.0,
            "presentations_limit": 1,
            "monthly_credits_limit": 1,
            "ai_features_enabled": False,
            "started_at": datetime.utcnow(),
            "expires_at": None,
            "auto_renew": False,
            "credits_remaining": 1,
            "credits_reset_date": datetime.utcnow(),
            "start_date": datetime.utcnow(),
            "end_date": None
        },
        "usage_stats":{
            "total_conversions":0,
            "this_month_conversions":0,
            "total_credits_used":0,
            "last_conversion_date": None
        },
        "created_at": datetime.utcnow(),
        "updated_at": datetime.utcnow(),
        "last_login":None
    }
    users_collection=get_collection("users")
    result=await users_collection.insert_one(user_doc)

    user_doc["_id"]=result.inserted_id
    return UserInDB(**user_doc)



async def get_user_by_email(email:str)->Optional[UserInDB]:
    """Retrieve user by email"""
    
    users_collection=get_collection("users")
    user_doc=await users_collection.find_one({"email":email})

    if user_doc:
        return UserInDB(**user_doc)
    return None

async def get_user_by_id(user_id:str)->Optional[UserInDB]:
    """Retrieve user by user ID"""
    
    users_collection=get_collection("users")
    user_doc=await users_collection.find_one({"_id":ObjectId(user_id)})

    if user_doc:
        return UserInDB(**user_doc)
    return None


async def authenticate_user(email:str,password:str)->Optional[UserInDB]:
    """Authenticate user credentials"""
    
    print(f"🔍 authenticate_user called:")
    print(f"   - Email: {email}")
    print(f"   - Password length: {len(password) if password else 0}")
    
    user=await get_user_by_email(email)
    if not user:
        print(f"   - User not found for email: {email}")
        return None
    
    print(f"   - User found: {user.name} (Active: {user.is_active})")
    print(f"   - Stored hash starts with: {user.password_hash[:20]}...")
    
    password_valid = verify_password(password,user.password_hash)
    print(f"   - Password verification: {'Valid' if password_valid else 'Invalid'}")
    
    if not password_valid:
        print(f"   - Authentication failed: Invalid password")
        return None
    
    user_collection=get_collection("users")
    await user_collection.update_one({"_id":ObjectId(str(user.id))},{"$set":{"last_login":datetime.utcnow()}})

    print(f"   - Authentication successful!")
    return await get_user_by_id(str(user.id))



async def update_user_profile(user_id:str,updates:Dict[str,Any])->Optional[UserInDB]:
    """Update user profile information"""
    updates["updated_at"]=datetime.utcnow()
    users_collection=get_collection("users")
    result=await users_collection.update_one({"_id":ObjectId(str(user_id))},{"$set":updates})
    if result.modified_count>0:
        return await get_user_by_id(str(user_id))
    return None


async def deduct_user_credits(user_id:str,credits:int)->bool:
    """Deduct credits from user's account"""
    user=await get_user_by_id(user_id)
    if not user:
        return False
    if not has_credits_remaning(user,credits):
        return False
    user_collection=get_collection("users")
    result=await user_collection.update_one(
        {"_id":ObjectId(str(user_id))},
        {"$inc":{"subscription.credits_remaining": -credits,"usage_stats.total_credits_used":credits},
         "$set":{"updated_at":datetime.utcnow()}}
    )
    return result.modified_count>0


async def upgrade_subscription(user_id:str,new_plan:SubscriptionPlan)->bool:
    """Upgrade user's subscription plan"""
    user=await get_user_by_id(user_id)
    if not user:
        return False
    plan_credits={
        SubscriptionPlan.FREE:1,
        SubscriptionPlan.BASIC:20,
        SubscriptionPlan.PRO:100,
        SubscriptionPlan.ENTERPRISE:500
    }
    
    plan_prices={
        SubscriptionPlan.FREE: 0.0,
        SubscriptionPlan.BASIC: 25.0,
        SubscriptionPlan.PRO: 50.0,
        SubscriptionPlan.ENTERPRISE: 100.0
    }

    users_collection=get_collection("users")
    result=await users_collection.update_one(
        {"_id":ObjectId(str(user_id))},
        {"$set":{
            "subscription.plan": new_plan,
            "subscription.status": SubscriptionStatus.ACTIVE,
            "subscription.price_per_month": plan_prices.get(new_plan, 0.0),
            "subscription.monthly_credits_limit": plan_credits.get(new_plan, 1),
            "subscription.presentations_limit": plan_credits.get(new_plan, 1),
            "subscription.started_at": datetime.utcnow(),
            "subscription.credits_remaining": plan_credits.get(new_plan, 1),
            "subscription.credits_reset_date": datetime.utcnow(),
            "updated_at": datetime.utcnow()
        }}
    )
    return result.modified_count>0
        


async def get_user_usage_stats(user_id:str)->Dict[str,Any]:
    """Get user's usage statistics"""
    user=await get_user_by_id(user_id)
    if not user:
        return {}
    return get_credit_usage_stats(user)
