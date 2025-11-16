"""
Quick Fix: Update AI_PRO Subscription Limits
=============================================
This script fixes the credit and presentation limits for users with ai_pro plan.
"""

import asyncio
from motor.motor_asyncio import AsyncIOMotorClient
import os
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()

MONGODB_URL = os.getenv("DATABASE_URL", "mongodb://localhost:27017")
DATABASE_NAME = os.getenv("DATABASE_NAME", "findeck_db")

async def fix_ai_pro_limits():
    print("="*80)
    print("🔧 FIXING AI_PRO SUBSCRIPTION LIMITS")
    print("="*80)
    
    client = AsyncIOMotorClient(MONGODB_URL)
    db = client[DATABASE_NAME]
    users_collection = db["users"]
    
    try:
        # Find all users with ai_pro plan
        ai_pro_users = await users_collection.find({"subscription.plan": "ai_pro"}).to_list(length=100)
        
        print(f"\n📊 Found {len(ai_pro_users)} AI_PRO users")
        
        if not ai_pro_users:
            print("✅ No AI_PRO users need fixing")
            return
        
        # Update each user
        for user in ai_pro_users:
            email = user.get("email", "Unknown")
            print(f"\n🔄 Updating: {email}")
            
            result = await users_collection.update_one(
                {"_id": user["_id"]},
                {
                    "$set": {
                        "subscription.monthly_credits_limit": 200,  # AI_PRO limit
                        "subscription.monthly_credits_used": 0,  # Reset used
                        "subscription.presentations_limit": 200,  # AI_PRO limit
                        "subscription.price_per_month": 99.0,  # AI_PRO price
                        "subscription.credits_remaining": 200,
                        "subscription.credits_reset_date": datetime.utcnow(),
                        "updated_at": datetime.utcnow()
                    }
                }
            )
            
            if result.modified_count > 0:
                print(f"   ✅ Updated successfully!")
                print(f"      • Credits: 200")
                print(f"      • Presentations: 200")
                print(f"      • Price: $99.00/month")
            else:
                print(f"   ⚠️  No changes made")
        
        print("\n" + "="*80)
        print("✅ ALL AI_PRO SUBSCRIPTIONS FIXED!")
        print("="*80)
        
    except Exception as e:
        print(f"\n❌ Error: {e}")
    finally:
        client.close()

if __name__ == "__main__":
    asyncio.run(fix_ai_pro_limits())
