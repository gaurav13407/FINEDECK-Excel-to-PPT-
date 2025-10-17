#!/usr/bin/env python3
"""
Comprehensive Storage System Test
Tests both LocalStorage and BackblazeB2Storage functionality
"""

import asyncio
import os
import tempfile
from datetime import datetime
from pathlib import Path
import sys

# Add the app directory to Python path for proper imports
app_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, app_dir)

# Set the module as main package to handle relative imports
if __name__ == "__main__":
    # For relative imports to work in the storage module
    import importlib.util
    
    # First import basic modules that don't have relative imports
    try:
        from storage.b2_storage import LocalStorage, BackblazeB2Storage, get_storage_client
        print("✓ Storage modules imported successfully")
    except ImportError as e:
        print(f"❌ Failed to import storage modules: {e}")
        sys.exit(1)
    
    # Try to import config settings
    try:
        from core.config import settings
        print("✓ Config settings imported successfully")
    except ImportError as e:
        print(f"⚠️  Config not available, using default values: {e}")
        # Create a simple settings object for testing
        class TestSettings:
            b2_application_key_id = os.getenv('B2_APPLICATION_KEY_ID')
            b2_application_key = os.getenv('B2_APPLICATION_KEY')
            b2_bucket_name = os.getenv('B2_BUCKET_NAME')
        settings = TestSettings()

class StorageTestSuite:
    """Test suite for storage systems"""
    
    def __init__(self):
        self.test_user_id = "test_user_12345"
        self.test_file_content = b"This is a test file for FinDeck storage system testing!"
        self.test_file_name = "test_document.txt"
        
    async def test_local_storage(self):
        """Test LocalStorage functionality"""
        print("\n🔵 Testing LocalStorage...")
        
        try:
            # Initialize Local Storage
            local_storage = LocalStorage()
            print("✓ LocalStorage initialized successfully")
            
            # Test file upload
            file_path = await local_storage.upload_file(
                self.test_file_content,
                self.test_file_name,
                self.test_user_id
            )
            print(f"✓ File uploaded successfully: {file_path}")
            
            # Test file download
            downloaded_content = await local_storage.download_file(file_path)
            if downloaded_content == self.test_file_content:
                print("✓ File download successful - content matches")
            else:
                print("❌ File download failed - content mismatch")
                return False
            
            # Test file existence
            exists = await local_storage.file_exists(file_path)
            if exists:
                print("✓ File existence check passed")
            else:
                print("❌ File existence check failed")
                return False
            
            # Test storage usage
            usage_stats = await local_storage.get_storage_usage(self.test_user_id)
            print(f"✓ Storage usage retrieved: {usage_stats}")
            
            # Test file deletion
            deleted = await local_storage.delete_file(file_path)
            if deleted:
                print("✓ File deletion successful")
            else:
                print("❌ File deletion failed")
                return False
            
            # Verify file is deleted
            exists_after_delete = await local_storage.file_exists(file_path)
            if not exists_after_delete:
                print("✓ File properly deleted - existence check passed")
            else:
                print("❌ File deletion verification failed")
                return False
            
            print("🟢 LocalStorage: ALL TESTS PASSED!")
            return True
            
        except Exception as e:
            print(f"❌ LocalStorage test failed: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    async def test_b2_storage(self):
        """Test BackblazeB2Storage functionality"""
        print("\n🔵 Testing BackblazeB2Storage...")
        
        # Check if B2 credentials are available
        if not all([
            settings.b2_application_key_id,
            settings.b2_application_key,
            settings.b2_bucket_name
        ]):
            print("⚠️  B2 credentials not configured - skipping B2 tests")
            print("   Set B2_APPLICATION_KEY_ID, B2_APPLICATION_KEY, and B2_BUCKET_NAME in .env")
            return True  # Not a failure, just skipped
        
        try:
            # Initialize B2 Storage
            b2_storage = BackblazeB2Storage(
                settings.b2_application_key_id,
                settings.b2_application_key,
                settings.b2_bucket_name
            )
            print("✓ BackblazeB2Storage initialized successfully")
            
            # Test file upload
            file_path = await b2_storage.upload_file(
                self.test_file_content,
                self.test_file_name,
                self.test_user_id
            )
            print(f"✓ File uploaded to B2: {file_path}")
            
            # Test file download
            downloaded_content = await b2_storage.download_file(file_path)
            if downloaded_content == self.test_file_content:
                print("✓ File download from B2 successful - content matches")
            else:
                print("❌ File download from B2 failed - content mismatch")
                return False
            
            # Test file existence
            exists = await b2_storage.file_exists(file_path)
            if exists:
                print("✓ B2 file existence check passed")
            else:
                print("❌ B2 file existence check failed")
                return False
            
            # Test storage usage
            usage_stats = await b2_storage.get_storage_usage(self.test_user_id)
            print(f"✓ B2 storage usage retrieved: {usage_stats}")
            
            # Test file deletion
            deleted = await b2_storage.delete_file(file_path)
            if deleted:
                print("✓ File deletion from B2 successful")
            else:
                print("❌ File deletion from B2 failed")
                return False
            
            # Verify file is deleted
            exists_after_delete = await b2_storage.file_exists(file_path)
            if not exists_after_delete:
                print("✓ B2 file properly deleted - existence check passed")
            else:
                print("❌ B2 file deletion verification failed")
                return False
            
            print("🟢 BackblazeB2Storage: ALL TESTS PASSED!")
            return True
            
        except Exception as e:
            print(f"❌ BackblazeB2Storage test failed: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    async def test_factory_function(self):
        """Test storage factory function"""
        print("\n🔵 Testing Storage Factory Function...")
        
        try:
            # Test default (should return LocalStorage)
            storage_client = get_storage_client()
            if isinstance(storage_client, LocalStorage):
                print("✓ Factory function returns LocalStorage by default")
            else:
                print("❌ Factory function default test failed")
                return False
            
            # Test B2 configuration
            if all([settings.b2_application_key_id, settings.b2_application_key, settings.b2_bucket_name]):
                b2_client = get_storage_client(
                    use_b2=True,
                    application_key_id=settings.b2_application_key_id,
                    application_key=settings.b2_application_key,
                    bucket_name=settings.b2_bucket_name
                )
                if isinstance(b2_client, BackblazeB2Storage):
                    print("✓ Factory function returns BackblazeB2Storage when configured")
                else:
                    print("❌ Factory function B2 test failed")
                    return False
            else:
                print("⚠️  B2 credentials not available - skipping B2 factory test")
            
            print("🟢 Storage Factory: ALL TESTS PASSED!")
            return True
            
        except Exception as e:
            print(f"❌ Storage factory test failed: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def print_test_environment(self):
        """Print test environment information"""
        print("📊 Test Environment Information:")
        print(f"   Python Version: {sys.version}")
        print(f"   Current Directory: {os.getcwd()}")
        print(f"   B2 Configured: {'Yes' if settings.b2_application_key_id else 'No'}")
        if settings.b2_bucket_name:
            print(f"   B2 Bucket: {settings.b2_bucket_name}")
        print()

async def main():
    """Run all storage tests"""
    print("🚀 FinDeck Storage System Test Suite")
    print("=" * 50)
    
    test_suite = StorageTestSuite()
    test_suite.print_test_environment()
    
    results = []
    
    # Test LocalStorage
    local_result = await test_suite.test_local_storage()
    results.append(("LocalStorage", local_result))
    
    # Test BackblazeB2Storage
    b2_result = await test_suite.test_b2_storage()
    results.append(("BackblazeB2Storage", b2_result))
    
    # Test Factory Function
    factory_result = await test_suite.test_factory_function()
    results.append(("Factory Function", factory_result))
    
    # Print Summary
    print("\n" + "=" * 50)
    print("📋 TEST SUMMARY")
    print("=" * 50)
    
    all_passed = True
    for test_name, result in results:
        status = "PASS" if result else "FAIL"
        emoji = "✅" if result else "❌"
        print(f"{emoji} {test_name}: {status}")
        if not result:
            all_passed = False
    
    print("\n" + "=" * 50)
    if all_passed:
        print("🎉 ALL STORAGE TESTS PASSED!")
        print("Your storage system is working correctly!")
    else:
        print("🚨 SOME TESTS FAILED!")
        print("Please check the error messages above.")
    print("=" * 50)
    
    return all_passed

if __name__ == "__main__":
    # Run the test suite
    success = asyncio.run(main())
    sys.exit(0 if success else 1)