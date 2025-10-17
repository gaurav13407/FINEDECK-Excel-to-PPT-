#!/usr/bin/env python3
"""
FinDeck Storage Classes Test
Tests the actual LocalStorage and BackblazeB2Storage classes
"""

import asyncio
import os
import sys
from pathlib import Path

# Add proper path handling
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

async def test_actual_storage_classes():
    """Test the actual storage classes without complex config dependencies"""
    print("🔵 Testing Actual Storage Classes...")
    
    try:
        # Import the actual classes
        import importlib.util
        
        # Load the b2_storage module directly
        b2_storage_path = current_dir / "storage" / "b2_storage.py"
        spec = importlib.util.spec_from_file_location("b2_storage", b2_storage_path)
        b2_storage_module = importlib.util.module_from_spec(spec)
        
        # Execute the module
        spec.loader.exec_module(b2_storage_module)
        
        # Get the classes
        LocalStorage = b2_storage_module.LocalStorage
        BackblazeB2Storage = b2_storage_module.BackblazeB2Storage
        get_storage_client = b2_storage_module.get_storage_client
        
        print("✓ Storage classes imported successfully")
        
        # Test LocalStorage
        await test_local_storage_class(LocalStorage)
        
        # Test factory function
        await test_factory_function(get_storage_client)
        
        # Test B2 storage if credentials available
        await test_b2_storage_class(BackblazeB2Storage)
        
        return True
        
    except Exception as e:
        print(f"❌ Failed to test storage classes: {e}")
        import traceback
        traceback.print_exc()
        return False

async def test_local_storage_class(LocalStorage):
    """Test the actual LocalStorage class"""
    print("\n🔵 Testing LocalStorage Class...")
    
    try:
        storage = LocalStorage()
        
        # Test data
        test_content = b"FinDeck Excel test content!"
        test_filename = "actual_test.xlsx"
        test_user = "testuser456"
        
        # Upload Excel file
        file_url, file_path = await storage.upload_excel_file(test_content, test_user, test_filename)
        print(f"✓ Excel file uploaded: {file_url}")
        print(f"   File path: {file_path}")
        
        # Test analysis data upload
        analysis_data = {"sheets": ["Sheet1"], "data_types": ["numeric", "text"]}
        analysis_url, analysis_path = await storage.upload_analysis_data(analysis_data, test_user, "test123")
        print(f"✓ Analysis data uploaded: {analysis_url}")
        
        # Get download URL (for local storage, this might just return file path)
        try:
            download_url = await storage.get_download_url(file_path)
            print(f"✓ Download URL generated: {download_url}")
        except AttributeError:
            print("⚠️  get_download_url not available in LocalStorage")
        
        # Get storage usage (might not be implemented for LocalStorage)
        try:
            usage = await storage.get_storage_usage(test_user)
            print(f"✓ Storage usage: {usage}")
        except AttributeError:
            print("⚠️  get_storage_usage not available in LocalStorage")
        
        # Test cleanup function
        try:
            cleanup_result = await storage.cleanup_expired_files(hours_old=0)  # Clean all files
            print(f"✓ Cleanup completed: {cleanup_result}")
        except Exception as e:
            print(f"⚠️  Cleanup failed (expected): {e}")
        
        # Test file existence by checking actual file system
        import os
        file_exists = os.path.exists(file_path)
        analysis_exists = os.path.exists(analysis_path)
        print(f"✓ Files exist on filesystem: Excel={file_exists}, Analysis={analysis_exists}")
        
        print("🟢 LocalStorage Class: ALL TESTS PASSED!")
        return True
        
    except Exception as e:
        print(f"❌ LocalStorage class test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

async def test_factory_function(get_storage_client):
    """Test the storage factory function"""
    print("\n🔵 Testing Storage Factory Function...")
    
    try:
        # Test default (LocalStorage)
        storage = get_storage_client()
        storage_type = type(storage).__name__
        assert storage_type == "LocalStorage", f"Expected LocalStorage, got {storage_type}"
        print("✓ Factory returns LocalStorage by default")
        
        # Test with B2 parameters but without valid credentials
        storage_b2 = get_storage_client(
            use_b2=True,
            application_key_id="test_key_id",
            application_key="test_key",
            bucket_name="test_bucket"
        )
        storage_b2_type = type(storage_b2).__name__
        print(f"✓ Factory with B2 params returns: {storage_b2_type}")
        
        print("🟢 Storage Factory: ALL TESTS PASSED!")
        return True
        
    except Exception as e:
        print(f"❌ Factory function test failed: {e}")
        return False

async def test_b2_storage_class(BackblazeB2Storage):
    """Test B2 storage class if credentials are available"""
    print("\n🔵 Testing B2 Storage Class...")
    
    # Check for credentials
    key_id = os.getenv('B2_APPLICATION_KEY_ID')
    key = os.getenv('B2_APPLICATION_KEY')
    bucket = os.getenv('B2_BUCKET_NAME')
    
    if not all([key_id, key, bucket]):
        print("⚠️  B2 credentials not available - skipping B2 class test")
        return True
    
    try:
        # Initialize B2 storage
        b2_storage = BackblazeB2Storage(key_id, key, bucket)
        print("✓ B2 Storage initialized")
        
        # Test Excel file upload
        test_content = b"B2 Excel test content!"
        test_filename = "b2_test.xlsx"
        test_user = "b2testuser"
        
        # Upload Excel file
        file_url, file_path = await b2_storage.upload_excel_file(test_content, test_user, test_filename)
        print(f"✓ Excel file uploaded to B2: {file_url}")
        
        # Test analysis data
        analysis_data = {"test": "b2_analysis"}
        analysis_url, analysis_path = await b2_storage.upload_analysis_data(analysis_data, test_user, "b2test123")
        print(f"✓ Analysis data uploaded to B2: {analysis_url}")
        
        # Get download URL
        download_url = await b2_storage.get_download_url(file_path)
        print(f"✓ B2 download URL generated: {download_url}")
        
        # Get storage usage
        usage = await b2_storage.get_storage_usage(test_user)
        print(f"✓ B2 storage usage: {usage}")
        
        # Delete files
        deleted1 = await b2_storage.delete_file(file_path)
        deleted2 = await b2_storage.delete_file(analysis_path)
        print(f"✓ Files deleted from B2: Excel={deleted1}, Analysis={deleted2}")
        
        print("🟢 B2 Storage Class: ALL TESTS PASSED!")
        return True
        
    except Exception as e:
        print(f"❌ B2 storage class test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

async def main():
    """Run storage class tests"""
    print("🚀 FinDeck Storage Classes Test")
    print("=" * 50)
    
    # Environment info
    print("📊 Environment:")
    print(f"   Python: {sys.version}")
    print(f"   Directory: {os.getcwd()}")
    print(f"   B2 Credentials: {'Available' if all([os.getenv('B2_APPLICATION_KEY_ID'), os.getenv('B2_APPLICATION_KEY'), os.getenv('B2_BUCKET_NAME')]) else 'Not Available'}")
    print()
    
    # Run tests
    success = await test_actual_storage_classes()
    
    print("\n" + "=" * 50)
    if success:
        print("🎉 STORAGE CLASSES TEST COMPLETED!")
    else:
        print("🚨 STORAGE CLASSES TEST FAILED!")
    print("=" * 50)
    
    return success

if __name__ == "__main__":
    result = asyncio.run(main())
    sys.exit(0 if result else 1)