"""
Database Migration Script: Normalize User Subscription Data
============================================================

This script updates all user documents to have canonical subscription limit values
derived from PLAN_CONFIGS instead of stale stored snapshots.

Usage:
    python scripts/migrate_subscriptions.py [--dry-run]

Options:
    --dry-run    Preview changes without modifying the database
"""

import asyncio
import sys
from pathlib import Path
from datetime import datetime

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src" / "backend"))

from motor.motor_asyncio import AsyncIOMotorClient
from app.core.config import settings
from app.models.user import PLAN_CONFIGS


async def migrate_subscriptions(dry_run: bool = False):
    """
    Migrate all user subscription documents to use canonical PLAN_CONFIGS values.
    
    Args:
        dry_run: If True, preview changes without modifying database
    """
    print(f"{'='*60}")
    print(f"Subscription Migration Script")
    print(f"{'='*60}")
    print(f"Mode: {'DRY RUN (no changes will be made)' if dry_run else 'LIVE (database will be updated)'}")
    print(f"Database: {settings.DATABASE_NAME}")
    print(f"Timestamp: {datetime.now().isoformat()}")
    print(f"{'='*60}\n")
    
    # Connect to MongoDB
    client = AsyncIOMotorClient(settings.DATABASE_URL)
    db = client[settings.DATABASE_NAME]
    
    try:
        # Fetch all users
        users_cursor = db.users.find({})
        users_list = await users_cursor.to_list(length=None)
        
        print(f"Found {len(users_list)} users to process\n")
        
        updated_count = 0
        skipped_count = 0
        error_count = 0
        
        for user in users_list:
            user_id = user.get('_id')
            email = user.get('email', 'unknown')
            subscription = user.get('subscription', {})
            plan = subscription.get('plan', 'free')
            
            # Get canonical limits from PLAN_CONFIGS
            plan_config = PLAN_CONFIGS.get(plan, PLAN_CONFIGS['free'])
            
            # Current stored values
            current_pres_limit = subscription.get('presentations_limit')
            current_credits_limit = subscription.get('credits_limit')
            current_file_limit = subscription.get('monthly_file_limit')
            
            # Canonical values from config
            canonical_pres_limit = plan_config['presentations_limit']
            canonical_credits_limit = plan_config['credits']
            canonical_file_limit = plan_config['monthly_file_limit']
            
            # Check if update needed
            needs_update = (
                current_pres_limit != canonical_pres_limit or
                current_credits_limit != canonical_credits_limit or
                current_file_limit != canonical_file_limit
            )
            
            if needs_update:
                print(f"User: {email} (Plan: {plan.upper()})")
                print(f"  presentations_limit: {current_pres_limit} → {canonical_pres_limit}")
                print(f"  credits_limit: {current_credits_limit} → {canonical_credits_limit}")
                print(f"  monthly_file_limit: {current_file_limit} → {canonical_file_limit}")
                
                if not dry_run:
                    try:
                        result = await db.users.update_one(
                            {'_id': user_id},
                            {'$set': {
                                'subscription.presentations_limit': canonical_pres_limit,
                                'subscription.credits_limit': canonical_credits_limit,
                                'subscription.monthly_file_limit': canonical_file_limit,
                                'subscription.updated_at': datetime.utcnow()
                            }}
                        )
                        
                        if result.modified_count > 0:
                            print(f"  ✅ Updated successfully")
                            updated_count += 1
                        else:
                            print(f"  ⚠️ No changes made (already up to date?)")
                            skipped_count += 1
                    except Exception as e:
                        print(f"  ❌ Error updating: {e}")
                        error_count += 1
                else:
                    print(f"  [DRY RUN] Would update")
                    updated_count += 1
                
                print()
            else:
                skipped_count += 1
        
        # Summary
        print(f"\n{'='*60}")
        print(f"Migration Summary")
        print(f"{'='*60}")
        print(f"Total users processed: {len(users_list)}")
        print(f"Users {'would be ' if dry_run else ''}updated: {updated_count}")
        print(f"Users skipped (already correct): {skipped_count}")
        if error_count > 0:
            print(f"Errors encountered: {error_count}")
        print(f"{'='*60}\n")
        
        if dry_run:
            print("✨ This was a DRY RUN. No changes were made to the database.")
            print("   Run without --dry-run to apply changes.\n")
        else:
            print("✅ Migration completed successfully!\n")
    
    except Exception as e:
        print(f"\n❌ Fatal error during migration: {e}")
        raise
    
    finally:
        client.close()


async def main():
    """Parse arguments and run migration."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Migrate user subscription data to canonical PLAN_CONFIGS values'
    )
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview changes without modifying the database'
    )
    
    args = parser.parse_args()
    
    # Run migration
    await migrate_subscriptions(dry_run=args.dry_run)


if __name__ == "__main__":
    asyncio.run(main())
