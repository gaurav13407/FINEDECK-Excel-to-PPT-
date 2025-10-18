#!/usr/bin/env python3
"""
Reset user password to a known value for testing
"""
import asyncio
import sys
import os

# Add the app directory to Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'app'))

from core.database import connect_to_mongo, get_collection
from core.security import get_password_hash
from bson import ObjectId

async def reset_password():
    print("🔄 Resetting user password...")
    
    try:
        # Connect to database
        await connect_to_mongo()
        
        # User to reset
        test_email = "gaurav13407@outlook.com"
        new_password = "password123"  # Known password for testing
        
        print(f"👤 User: {test_email}")
        print(f"🔑 New password: {new_password}")
        
        # Generate new hash using current method
        new_hash = get_password_hash(new_password)
        print(f"🔐 New hash: {new_hash}")
        print(f"🔐 Hash length: {len(new_hash)}")
        
        # Update in database
        users_collection = get_collection("users")
        
        result = await users_collection.update_one(
            {"email": test_email},
            {
                "$set": {
                    "password_hash": new_hash,
                    "updated_at": "2025-10-18T10:30:00.000Z"
                }
            }
        )
        
        if result.modified_count > 0:
            print(f"✅ Password updated successfully!")
            print(f"📧 Email: {test_email}")
            print(f"🔑 Password: {new_password}")
            print(f"\n🧪 Now try logging in with these credentials!")
        else:
            print(f"❌ Failed to update password (user not found?)")
        
    except Exception as e:
        print(f"❌ Error during password reset: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(reset_password())