#!/usr/bin/env python3
"""
Test different password hashing methods to find the correct one
"""
import asyncio
import sys
import os
import hashlib

# Add the app directory to Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'app'))

from core.database import connect_to_mongo, get_collection
from services.user_service import get_user_by_email

async def test_password_hashing():
    print("🔍 Testing password hashing methods...")
    
    try:
        # Connect to database
        await connect_to_mongo()
        
        # Get the user
        test_email = "gaurav13407@outlook.com"
        user = await get_user_by_email(test_email)
        
        if not user:
            print("❌ User not found")
            return
        
        stored_hash = user.password_hash
        print(f"🔑 Stored hash: {stored_hash}")
        print(f"🔑 Hash length: {len(stored_hash)}")
        
        # Test different common passwords with different hashing methods
        test_passwords = [
            "password", "123456", "admin", "test", "findeck", 
            "gaurav", "123", "password123", "admin123", 
            "findeck123", "Gaurav123", "gaurav123"
        ]
        
        print(f"\n🧪 Testing common passwords with different hash methods:")
        
        for password in test_passwords:
            # Method 1: Plain MD5
            md5_hash = hashlib.md5(password.encode()).hexdigest()
            if md5_hash == stored_hash:
                print(f"✅ MATCH! Password: '{password}' using MD5")
                return password
            
            # Method 2: Plain SHA256
            sha256_hash = hashlib.sha256(password.encode()).hexdigest()
            if sha256_hash == stored_hash:
                print(f"✅ MATCH! Password: '{password}' using SHA256")
                return password
            
            # Method 3: SHA1
            sha1_hash = hashlib.sha1(password.encode()).hexdigest()
            if sha1_hash == stored_hash:
                print(f"✅ MATCH! Password: '{password}' using SHA1")
                return password
        
        print("❌ No matches found with common passwords and hash methods")
        print("💡 The password might be:")
        print("   1. A custom password you chose")
        print("   2. Using a different salt or hash method")
        print("   3. Need to reset the password")
        
        # Offer to reset password
        print(f"\n🔄 Would you like to reset the password for {test_email}?")
        print("   We can set it to a known password like 'password123'")
        
    except Exception as e:
        print(f"❌ Error during test: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(test_password_hashing())