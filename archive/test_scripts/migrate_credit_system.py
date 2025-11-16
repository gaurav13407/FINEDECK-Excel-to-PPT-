"""
Database Migration Script - Align Credits with PPT Limits
Fixes the credit system to match PPT limits (1 credit = 1 PPT)
"""

import asyncio
from motor.motor_asyncio import AsyncIOMotorClient
from datetime import datetime
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# MongoDB connection
MONGODB_URL = os.getenv("MONGODB_URL", "mongodb://localhost:27017")
DATABASE_NAME = os.getenv("DATABASE_NAME", "findeck")

async def migrate_credit_system():
    """Align credit limits with PPT limits for all users"""
    
    print("=" * 80)
    print("CREDIT SYSTEM MIGRATION")
    print("=" * 80)
    print("\nAligning credit limits with PPT limits (1 credit = 1 PPT)\n")
    
    # Connect to MongoDB
    client = AsyncIOMotorClient(MONGODB_URL)
    db = client[DATABASE_NAME]
    users_collection = db['users']
    
    # Define correct credit limits per tier
    tier_limits = {
        'free': {'ppt_limit': 1, 'credit_limit': 1},
        'basic': {'ppt_limit': 7, 'credit_limit': 7},  # Was 10, now 7
        'pro': {'ppt_limit': 15, 'credit_limit': 15},  # Was 50, now 15
        'ai_pro': {'ppt_limit': -1, 'credit_limit': -1},  # Was 1000, now -1 (unlimited)
        'enterprise': {'ppt_limit': 1000, 'credit_limit': 1000}
    }
    
    print("📊 Credit Limit Corrections:")
    for tier, limits in tier_limits.items():
        print(f"   {tier.upper():12s} → {limits['credit_limit']:4d} credits (PPT limit: {limits['ppt_limit']})")
    print()
    
    # Count users per tier
    print("📈 Analyzing current users...")
    for tier, limits in tier_limits.items():
        count = await users_collection.count_documents({"subscription.plan": tier})
        if count > 0:
            print(f"   {tier.upper():12s}: {count} users")
    print()
    
    # Update each tier
    total_updated = 0
    
    for tier, limits in tier_limits.items():
        print(f"🔧 Updating {tier.upper()} tier users...")
        
        result = await users_collection.update_many(
            {"subscription.plan": tier},
            {
                "$set": {
                    "subscription.monthly_credits_limit": limits['credit_limit'],
                    "subscription.presentations_limit": limits['ppt_limit'],
                    "updated_at": datetime.utcnow()
                }
            }
        )
        
        if result.modified_count > 0:
            print(f"   ✅ Updated {result.modified_count} users")
            total_updated += result.modified_count
        else:
            print(f"   ⏭️  No users to update")
    
    print()
    print("=" * 80)
    print(f"✅ MIGRATION COMPLETE! Updated {total_updated} users")
    print("=" * 80)
    
    # Show summary of current state
    print("\n📊 Current System State:")
    print(f"{'Tier':<12} {'PPT Limit':>12} {'Credit Limit':>14} {'Users':>8}")
    print("-" * 50)
    
    for tier, limits in tier_limits.items():
        count = await users_collection.count_documents({"subscription.plan": tier})
        ppt_display = "Unlimited" if limits['ppt_limit'] == -1 else str(limits['ppt_limit'])
        credit_display = "Unlimited" if limits['credit_limit'] == -1 else str(limits['credit_limit'])
        print(f"{tier:<12} {ppt_display:>12} {credit_display:>14} {count:>8}")
    
    # Close connection
    client.close()
    print("\n✅ All done! Credit system is now aligned with PPT limits.")


async def verify_migration():
    """Verify that all users have aligned credits"""
    
    print("\n" + "=" * 80)
    print("VERIFICATION CHECK")
    print("=" * 80)
    
    # Connect to MongoDB
    client = AsyncIOMotorClient(MONGODB_URL)
    db = client[DATABASE_NAME]
    users_collection = db['users']
    
    # Check for misaligned users
    misaligned_users = []
    
    async for user in users_collection.find():
        plan = user.get('subscription', {}).get('plan', 'free')
        credit_limit = user.get('subscription', {}).get('monthly_credits_limit', 0)
        ppt_limit = user.get('subscription', {}).get('presentations_limit', 0)
        
        # Expected alignment
        expected = {
            'free': 1,
            'basic': 7,
            'pro': 15,
            'ai_pro': -1,
            'enterprise': 1000
        }
        
        if credit_limit != expected.get(plan, 1) or credit_limit != ppt_limit:
            misaligned_users.append({
                'email': user.get('email'),
                'plan': plan,
                'credit_limit': credit_limit,
                'ppt_limit': ppt_limit,
                'expected': expected.get(plan, 1)
            })
    
    if misaligned_users:
        print("\n⚠️  Found misaligned users:")
        for user in misaligned_users:
            print(f"   {user['email']:30s} | {user['plan']:10s} | "
                  f"Credits: {user['credit_limit']:4d} | "
                  f"PPTs: {user['ppt_limit']:4d} | "
                  f"Expected: {user['expected']:4d}")
        print(f"\n❌ {len(misaligned_users)} users still have misaligned limits!")
    else:
        print("\n✅ All users have properly aligned credit and PPT limits!")
    
    client.close()


async def reset_monthly_usage():
    """Reset monthly usage counters (run at start of each month)"""
    
    print("\n" + "=" * 80)
    print("RESET MONTHLY USAGE")
    print("=" * 80)
    
    # Connect to MongoDB
    client = AsyncIOMotorClient(MONGODB_URL)
    db = client[DATABASE_NAME]
    users_collection = db['users']
    
    current_month = datetime.utcnow().strftime("%Y-%m")
    
    result = await users_collection.update_many(
        {},
        {
            "$set": {
                "subscription.monthly_credits_used": 0,
                "usage_stats.this_month_conversions": 0,
                "usage_stats.last_reset_month": current_month,
                "updated_at": datetime.utcnow()
            }
        }
    )
    
    print(f"✅ Reset usage for {result.modified_count} users")
    print(f"   Month: {current_month}")
    
    client.close()


async def main():
    """Main migration function"""
    
    print("\n🚀 Starting Credit System Migration...\n")
    
    # Run migration
    await migrate_credit_system()
    
    # Verify
    await verify_migration()
    
    # Ask about monthly reset
    print("\n" + "=" * 80)
    print("OPTIONAL: Monthly Usage Reset")
    print("=" * 80)
    response = input("\nDo you want to reset monthly usage counters? (y/n): ")
    if response.lower() == 'y':
        await reset_monthly_usage()
    
    print("\n🎉 Migration complete!")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\n⚠️  Migration cancelled by user")
    except Exception as e:
        print(f"\n\n❌ Migration failed: {str(e)}")
        import traceback
        traceback.print_exc()
