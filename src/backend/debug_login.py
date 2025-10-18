#!/usr/bin/env python3
"""
Debug login issues - check user credentials and authentication flow
"""
import asyncio
import sys
import os

# Add the app directory to Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'app'))

from core.database import connect_to_mongo, get_collection
from services.user_service import get_user_by_email, authenticate_user
from core.security import verify_password

async def debug_login():
    print("🔍 Debugging login process...")
    
    try:
        # Connect to database
        await connect_to_mongo()
        print("✅ Database connected")
        
        # Get email to test
        test_email = input("📧 Enter email to test: ").strip()
        test_password = input("🔑 Enter password to test: ").strip()
        
        if not test_email or not test_password:
            print("❌ Email and password required")
            return
        
        print(f"\n🔍 Step 1: Checking if user exists...")
        user = await get_user_by_email(test_email)
        
        if not user:
            print(f"❌ User not found for email: {test_email}")
            
            # Check if it's a case sensitivity issue
            users_collection = get_collection("users")
            user_doc = await users_collection.find_one({"email": {"$regex": f"^{test_email}$", "$options": "i"}})
            
            if user_doc:
                print(f"⚠️  Found user with different case: {user_doc.get('email')}")
                print("💡 Email case sensitivity might be the issue")
            else:
                print("💡 User definitely doesn't exist in database")
            
            return
        
        print(f"✅ User found: {user.name} ({user.email})")
        print(f"   - Active: {user.is_active}")
        print(f"   - Has password hash: {'Yes' if user.password_hash else 'No'}")
        
        print(f"\n🔍 Step 2: Testing password verification...")
        
        # Test password directly
        password_valid = verify_password(test_password, user.password_hash)
        print(f"   - Password verification result: {password_valid}")
        
        if not password_valid:
            print("❌ Password doesn't match stored hash")
            print("💡 This is why authentication is failing")
            
            # Let's check what the hash looks like
            print(f"   - Stored hash starts with: {user.password_hash[:20]}...")
            print(f"   - Hash algorithm appears to be: {'bcrypt' if user.password_hash.startswith('$2b$') else 'unknown'}")
        else:
            print("✅ Password matches!")
        
        print(f"\n🔍 Step 3: Testing full authentication flow...")
        auth_result = await authenticate_user(test_email, test_password)
        
        if auth_result:
            print("✅ Full authentication successful!")
            print(f"   - Authenticated user: {auth_result.name}")
        else:
            print("❌ Full authentication failed")
        
        # Additional debugging info
        print(f"\n📊 Debug Summary:")
        print(f"   - User exists: {'Yes' if user else 'No'}")
        print(f"   - User active: {user.is_active if user else 'N/A'}")
        print(f"   - Password valid: {password_valid if user else 'N/A'}")
        print(f"   - Auth successful: {'Yes' if auth_result else 'No'}")
        
    except Exception as e:
        print(f"❌ Error during debug: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(debug_login())