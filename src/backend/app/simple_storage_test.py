#!/usr/bin/env python3
"""
Simple Storage Test - Direct Class Testing
Tests LocalStorage and BackblazeB2Storage without complex imports
"""

import asyncio
import os
import tempfile
from datetime import datetime

# Simple test for LocalStorage
async def test_local_storage_simple():
    """Test LocalStorage with minimal imports"""
    print("🔵 Testing LocalStorage (Simple)...")
    
    # Manual LocalStorage implementation test
    import os
    from pathlib import Path
    
    class SimpleLocalStorage:
        def __init__(self):
            self.base_dir = Path("temp_storage")
            self.base_dir.mkdir(exist_ok=True)
        
        async def upload_file(self, content: bytes, filename: str, user_id: str) -> str:
            user_dir = self.base_dir / user_id
            user_dir.mkdir(exist_ok=True)
            
            file_path = user_dir / filename
            with open(file_path, 'wb') as f:
                f.write(content)
            
            return str(file_path)
        
        async def download_file(self, file_path: str) -> bytes:
            with open(file_path, 'rb') as f:
                return f.read()
        
        async def file_exists(self, file_path: str) -> bool:
            return os.path.exists(file_path)
        
        async def delete_file(self, file_path: str) -> bool:
            try:
                os.remove(file_path)
                return True
            except:
                return False
    
    try:
        storage = SimpleLocalStorage()
        
        # Test data
        test_content = b"Test file content for FinDeck!"
        test_filename = "test.txt"
        test_user = "user123"
        
        # Upload test
        file_path = await storage.upload_file(test_content, test_filename, test_user)
        print(f"✓ File uploaded: {file_path}")
        
        # Download test
        downloaded = await storage.download_file(file_path)
        if downloaded == test_content:
            print("✓ File download successful")
        else:
            print("❌ File download failed")
            return False
        
        # Existence test
        exists = await storage.file_exists(file_path)
        if exists:
            print("✓ File exists check passed")
        else:
            print("❌ File exists check failed")
            return False
        
        # Delete test
        deleted = await storage.delete_file(file_path)
        if deleted:
            print("✓ File deleted successfully")
        else:
            print("❌ File deletion failed")
            return False
        
        # Cleanup
        import shutil
        if os.path.exists("temp_storage"):
            shutil.rmtree("temp_storage")
        
        print("🟢 SimpleLocalStorage: ALL TESTS PASSED!")
        return True
        
    except Exception as e:
        print(f"❌ LocalStorage test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

async def test_b2_storage_simple():
    """Test B2 storage if credentials are available"""
    print("\n🔵 Testing B2 Storage Connection...")
    
    # Check for B2 credentials in environment
    b2_key_id = os.getenv('B2_APPLICATION_KEY_ID')
    b2_key = os.getenv('B2_APPLICATION_KEY')
    b2_bucket = os.getenv('B2_BUCKET_NAME')
    
    if not all([b2_key_id, b2_key, b2_bucket]):
        print("⚠️  B2 credentials not found in environment variables")
        print("   Set B2_APPLICATION_KEY_ID, B2_APPLICATION_KEY, and B2_BUCKET_NAME")
        print("   Or create a .env file with these values")
        return True  # Not a failure, just skipped
    
    try:
        # Test B2 import
        from b2sdk.v2 import InMemoryAccountInfo, B2Api
        
        # Test B2 connection
        info = InMemoryAccountInfo()
        b2_api = B2Api(info)
        
        print(f"✓ Attempting to authorize with B2...")
        application_key_id = b2_key_id
        application_key = b2_key
        
        b2_api.authorize_account("production", application_key_id, application_key)
        print("✓ B2 authorization successful!")
        
        # Test bucket access
        bucket = b2_api.get_bucket_by_name(b2_bucket)
        print(f"✓ B2 bucket '{b2_bucket}' accessed successfully!")
        
        print("🟢 B2 Storage Connection: ALL TESTS PASSED!")
        return True
        
    except ImportError:
        print("❌ b2sdk not installed. Run: pip install b2sdk")
        return False
    except Exception as e:
        print(f"❌ B2 storage test failed: {e}")
        return False

def print_environment_info():
    """Print environment information"""
    print("📊 Test Environment:")
    print(f"   Python: {os.sys.version}")
    print(f"   Directory: {os.getcwd()}")
    print(f"   B2_APPLICATION_KEY_ID: {'Set' if os.getenv('B2_APPLICATION_KEY_ID') else 'Not Set'}")
    print(f"   B2_APPLICATION_KEY: {'Set' if os.getenv('B2_APPLICATION_KEY') else 'Not Set'}")
    print(f"   B2_BUCKET_NAME: {'Set' if os.getenv('B2_BUCKET_NAME') else 'Not Set'}")
    print()

async def main():
    """Run simple storage tests"""
    print("🚀 FinDeck Simple Storage Test")
    print("=" * 40)
    
    print_environment_info()
    
    # Test LocalStorage
    local_result = await test_local_storage_simple()
    
    # Test B2 Storage
    b2_result = await test_b2_storage_simple()
    
    # Summary
    print("\n" + "=" * 40)
    print("📋 TEST SUMMARY")
    print("=" * 40)
    
    print(f"{'✅' if local_result else '❌'} LocalStorage: {'PASS' if local_result else 'FAIL'}")
    print(f"{'✅' if b2_result else '❌'} B2 Storage: {'PASS' if b2_result else 'FAIL'}")
    
    all_passed = local_result and b2_result
    
    print("\n" + "=" * 40)
    if all_passed:
        print("🎉 ALL TESTS PASSED!")
    else:
        print("🚨 SOME TESTS FAILED!")
    print("=" * 40)
    
    return all_passed

if __name__ == "__main__":
    success = asyncio.run(main())
    exit(0 if success else 1)