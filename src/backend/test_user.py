#!/usr/bin/env python3
"""
Test script to check users in database and test authentication
"""
import asyncio
import sys
import os

# Add the app directory to Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'app'))

from core.database import get_collection
from services.user_service import get_user_by_email, authenticate_user
from core.security import get_password_hash

async def main():
    print("🔍 Checking database connection and users...")
    
    try:
        # Check if environment variables are loaded
        from core.config import settings
        print(f"📋 Current DATABASE_URL: {settings.database_url}")
        print(f"📋 Current DATABASE_NAME: {settings.database_name}")
        
        # Connect to database first
        from core.database import connect_to_mongo
        await connect_to_mongo()
        print("✅ Database connected successfully!")
        
        # Get users collection
        users_collection = get_collection("users")
        
        # Count total users
        user_count = await users_collection.count_documents({})
        print(f"📊 Total users in database: {user_count}")
        
        if user_count > 0:
            # List all users (without passwords)
            print("\n👥 Users in database:")
            async for user in users_collection.find({}, {"password_hash": 0}):
                print(f"   - Email: {user.get('email', 'N/A')}")
                print(f"     Name: {user.get('name', 'N/A')}")
                print(f"     Active: {user.get('is_active', 'N/A')}")
                print(f"     Created: {user.get('created_at', 'N/A')}")
                print()
        
        # Test a specific email (replace with the email you're trying to login with)
        test_email = input("\n📧 Enter the email you're trying to login with: ").strip()
        if test_email:
            print(f"\n🔍 Checking user: {test_email}")
            user = await get_user_by_email(test_email)
            
            if user:
                print("✅ User found in database!")
                print(f"   - Name: {user.name}")
                print(f"   - Active: {user.is_active}")
                print(f"   - Has password hash: {'Yes' if user.password_hash else 'No'}")
                
                # Test password
                test_password = input("\n🔑 Enter password to test: ").strip()
                if test_password:
                    auth_result = await authenticate_user(test_email, test_password)
                    if auth_result:
                        print("✅ Authentication successful!")
                    else:
                        print("❌ Authentication failed! Password doesn't match.")
            else:
                print("❌ User not found in database!")
                print("\n💡 You need to register this user first.")
                
                # Offer to create a test user
                create_user = input("\n❓ Create a test user? (y/n): ").lower() == 'y'
                if create_user:
                    name = input("Enter name: ").strip()
                    password = input("Enter password: ").strip()
                    
                    if name and password:
                        # Create user document
                        from datetime import datetime
                        from bson import ObjectId
                        
                        user_doc = {
                            "_id": ObjectId(),
                            "name": name,
                            "email": test_email,
                            "password_hash": get_password_hash(password),
                            "is_active": True,
                            "is_verified": True,
                            "created_at": datetime.utcnow(),
                            "updated_at": datetime.utcnow(),
                            "subscription": {
                                "plan": "free",
                                "status": "active",
                                "monthly_credits": 10,
                                "monthly_credits_used": 0,
                                "reset_date": datetime.utcnow().replace(day=1)
                            }
                        }
                        
                        result = await users_collection.insert_one(user_doc)
                        print(f"✅ User created with ID: {result.inserted_id}")
                        print("🎉 You can now try logging in!")
    
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(main())