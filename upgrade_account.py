#!/usr/bin/env python3
"""
Quick script to upgrade user account to premium plan
Usage: python upgrade_account.py <email> <plan>
Plans: basic, pro, enterprise
"""

import asyncio
import sys
from datetime import datetime, UTC
from motor.motor_asyncio import AsyncIOMotorClient
from bson import ObjectId

# MongoDB connection
MONGODB_URL = "mongodb+srv://FinDeck_DB:findeck127@cluster0.dlqmc.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"

async def upgrade_user_subscription(email: str, plan: str):
    """Upgrade user subscription plan"""
    
    # Validate plan
    valid_plans = ["basic", "pro", "enterprise"]
    if plan.lower() not in valid_plans:
        print(f"❌ Invalid plan. Choose from: {valid_plans}")
        return False
    
    try:
        # Connect to database
        client = AsyncIOMotorClient(MONGODB_URL)
        db = client.findeck_db
        users_collection = db.users
        
        # Find user by email
        user = await users_collection.find_one({"email": email})
        if not user:
            print(f"❌ User with email {email} not found")
            return False
        
        # Update subscription
        result = await users_collection.update_one(
            {"email": email},
            {
                "$set": {
                    "subscription.plan": plan.lower(),
                    "subscription.status": "active",
                    "subscription.monthly_credits_used": 0,  # Reset credits
                    "subscription.monthly_presentations_created": 0,  # Reset presentations
                    "subscription.updated_at": datetime.now(UTC)
                }
            }
        )
        
        if result.modified_count > 0:
            print(f"✅ Successfully upgraded {email} to {plan.upper()} plan!")
            print(f"🎉 Monthly limits reset!")
            return True
        else:
            print(f"❌ Failed to update user subscription")
            return False
            
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return False
    finally:
        client.close()

async def main():
    if len(sys.argv) != 3:
        print("Usage: python upgrade_account.py <email> <plan>")
        print("Plans: basic, pro, enterprise")
        print("Example: python upgrade_account.py user@email.com pro")
        return
    
    email = sys.argv[1]
    plan = sys.argv[2]
    
    print(f"🔄 Upgrading {email} to {plan.upper()} plan...")
    success = await upgrade_user_subscription(email, plan)
    
    if success:
        print("\n🎯 Next steps:")
        print("1. Refresh your browser page")
        print("2. Log out and log back in")
        print("3. Try uploading and converting files again!")
    else:
        print("\n❌ Upgrade failed. Check the error messages above.")

if __name__ == "__main__":
    asyncio.run(main())