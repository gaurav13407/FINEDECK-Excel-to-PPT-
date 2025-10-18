#!/usr/bin/env python3
"""
Non-interactive debug for specific user
"""
import asyncio
import sys
import os

# Add the app directory to Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'app'))

from core.database import connect_to_mongo, get_collection
from services.user_service import get_user_by_email, authenticate_user
from core.security import verify_password

async def debug_specific_user():
    print("🔍 Debugging login for specific user...")
    
    try:
        # Connect to database
        await connect_to_mongo()
        print("✅ Database connected")
        
        # Test with the most recent user
        test_email = "gaurav13407@outlook.com"  # One of the users we saw
        
        print(f"\n🔍 Testing user: {test_email}")
        print(f"🔍 Step 1: Checking if user exists...")
        
        user = await get_user_by_email(test_email)
        
        if not user:
            print(f"❌ User not found for email: {test_email}")
            
            # Let's list all users to see exact emails
            print("\n📋 All users in database:")
            users_collection = get_collection("users")
            async for user_doc in users_collection.find({}, {"email": 1, "name": 1, "_id": 0}):
                print(f"   - {user_doc.get('email')} ({user_doc.get('name')})")
            
            return
        
        print(f"✅ User found: {user.name} ({user.email})")
        print(f"   - Active: {user.is_active}")
        print(f"   - Has password hash: {'Yes' if user.password_hash else 'No'}")
        print(f"   - Password hash starts with: {user.password_hash[:30]}...")
        
        # Check what password hashing algorithm was used
        if user.password_hash.startswith('$2b$'):
            print("   - Hash algorithm: bcrypt")
        elif user.password_hash.startswith('pbkdf2'):
            print("   - Hash algorithm: PBKDF2")
        else:
            print(f"   - Hash algorithm: Unknown (starts with {user.password_hash[:10]})")
        
        print(f"\n💡 To test login, you need to provide the correct password for this user.")
        print(f"💡 If you don't remember the password, we can reset it.")
        
    except Exception as e:
        print(f"❌ Error during debug: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(debug_specific_user())