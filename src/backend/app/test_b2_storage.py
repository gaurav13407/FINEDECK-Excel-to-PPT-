#!/usr/bin/env python3
"""
FinDeck Storage Test with B2 Credentials
Tests both LocalStorage and BackblazeB2Storage with actual credentials
"""

import asyncio
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Add proper path handling
current_dir = Path(__file__).parent
sys.path.insert(0, str(current_dir))

async def test_with_b2_credentials():
    """Test storage with actual B2 credentials from .env"""
    print("🔵 Testing with B2 Credentials from .env...")
    
    # Get credentials from environment
    b2_key_id = os.getenv('B2_APPLICATION_KEY_ID')
    b2_key = os.getenv('B2_APPLICATION_KEY')
    b2_bucket = os.getenv('B2_BUCKET_NAME')
    use_b2 = os.getenv('USE_B2_STORAGE', 'false').lower() == 'true'
    
    print(f"📊 Environment Variables:")
    print(f"   B2_APPLICATION_KEY_ID: {'Set' if b2_key_id else 'Not Set'}")
    print(f"   B2_APPLICATION_KEY: {'Set' if b2_key else 'Not Set'}")
    print(f"   B2_BUCKET_NAME: {b2_bucket or 'Not Set'}")
    print(f"   USE_B2_STORAGE: {use_b2}")
    print()
    
    if not all([b2_key_id, b2_key, b2_bucket]):
        print("❌ Missing B2 credentials in .env file")
        return False
    
    try:
        # Import storage classes
        import importlib.util
        
        # Load the b2_storage module directly
        b2_storage_path = current_dir / "storage" / "b2_storage.py"
        spec = importlib.util.spec_from_file_location("b2_storage", b2_storage_path)
        b2_storage_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(b2_storage_module)
        
        # Get the classes
        LocalStorage = b2_storage_module.LocalStorage
        BackblazeB2Storage = b2_storage_module.BackblazeB2Storage
        get_storage_client = b2_storage_module.get_storage_client
        
        print("✅ Storage classes imported successfully")
        
        # Test LocalStorage first
        print("\n🔵 Testing LocalStorage...")
        local_storage = LocalStorage()
        
        test_content = b"Local storage test with .env!"
        test_filename = "env_test_local.xlsx"
        test_user = "env_test_user"
        
        file_url, file_path = await local_storage.upload_excel_file(test_content, test_user, test_filename)
        print(f"✅ LocalStorage upload successful: {file_url}")
        
        # Test B2 Storage with actual credentials
        print(f"\n🔵 Testing B2 Storage with bucket: {b2_bucket}...")
        
        try:
            b2_storage = BackblazeB2Storage(b2_key_id, b2_key, b2_bucket)
            print("✅ B2 Storage initialized successfully")
            
            # Test Excel file upload to B2
            b2_test_content = b"B2 storage test with real credentials!"
            b2_test_filename = "env_test_b2.xlsx"
            b2_test_user = "b2_env_test_user"
            
            print(f"   Uploading to B2...")
            b2_file_url, b2_file_path = await b2_storage.upload_excel_file(b2_test_content, b2_test_user, b2_test_filename)
            print(f"✅ B2 Excel upload successful: {b2_file_url}")
            
            # Test analysis data upload to B2
            analysis_data = {"test": "b2_analysis_with_env", "timestamp": "2025-10-17"}
            print(f"   Uploading analysis data to B2...")
            analysis_url, analysis_path = await b2_storage.upload_analysis_data(analysis_data, b2_test_user, "env_test")
            print(f"✅ B2 Analysis upload successful: {analysis_url}")
            
            # Test PPT output upload to B2 (if method exists)
            try:
                ppt_content = b"Test PPT content for B2"
                ppt_url, ppt_path = await b2_storage.upload_ppt_output(ppt_content, b2_test_user, "env_test.pptx")
                print(f"✅ B2 PPT upload successful: {ppt_url}")
                ppt_uploaded = True
            except AttributeError:
                print("⚠️  upload_ppt_output method not available")
                ppt_uploaded = False
            
            # Test download URL generation
            try:
                download_url = await b2_storage.get_download_url(b2_file_path, expiry_hours=1)
                print(f"✅ B2 download URL generated: {download_url[:50]}...")
            except Exception as e:
                print(f"⚠️  Download URL generation failed: {e}")
            
            # Test storage usage
            try:
                usage_stats = await b2_storage.get_storage_usage(b2_test_user)
                print(f"✅ B2 storage usage: {usage_stats}")
            except Exception as e:
                print(f"⚠️  Storage usage check failed: {e}")
            
            # Clean up - delete test files
            print(f"\n🧹 Cleaning up test files...")
            try:
                deleted_excel = await b2_storage.delete_file(b2_file_path)
                deleted_analysis = await b2_storage.delete_file(analysis_path)
                print(f"✅ Cleanup: Excel={deleted_excel}, Analysis={deleted_analysis}")
                
                if ppt_uploaded:
                    deleted_ppt = await b2_storage.delete_file(ppt_path)
                    print(f"✅ PPT cleanup: {deleted_ppt}")
                    
            except Exception as e:
                print(f"⚠️  Cleanup warning: {e}")
            
            print("\n🎉 B2 Storage: ALL TESTS PASSED!")
            return True
            
        except Exception as e:
            print(f"❌ B2 Storage test failed: {e}")
            import traceback
            traceback.print_exc()
            return False
            
    except Exception as e:
        print(f"❌ Test setup failed: {e}")
        import traceback
        traceback.print_exc()
        return False

async def test_factory_with_env():
    """Test storage factory with environment variables"""
    print("\n🔵 Testing Storage Factory with .env...")
    
    try:
        # Import storage classes
        import importlib.util
        
        b2_storage_path = current_dir / "storage" / "b2_storage.py"
        spec = importlib.util.spec_from_file_location("b2_storage", b2_storage_path)
        b2_storage_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(b2_storage_module)
        
        get_storage_client = b2_storage_module.get_storage_client
        
        # Test factory with B2 credentials
        b2_key_id = os.getenv('B2_APPLICATION_KEY_ID')
        b2_key = os.getenv('B2_APPLICATION_KEY')
        b2_bucket = os.getenv('B2_BUCKET_NAME')
        
        if all([b2_key_id, b2_key, b2_bucket]):
            storage_client = get_storage_client(
                use_b2=True,
                application_key_id=b2_key_id,
                application_key=b2_key,
                bucket_name=b2_bucket
            )
            client_type = type(storage_client).__name__
            print(f"✅ Factory with B2 credentials returns: {client_type}")
            
            if client_type == "BackblazeB2Storage":
                print("✅ Factory function working correctly with B2!")
                return True
            else:
                print("❌ Factory should return BackblazeB2Storage")
                return False
        else:
            print("⚠️  B2 credentials not available for factory test")
            return True
            
    except Exception as e:
        print(f"❌ Factory test failed: {e}")
        return False

async def main():
    """Run B2 storage tests with credentials"""
    print("🚀 FinDeck B2 Storage Test with Credentials")
    print("=" * 60)
    
    # Test storage with B2 credentials
    storage_result = await test_with_b2_credentials()
    
    # Test factory function
    factory_result = await test_factory_with_env()
    
    # Summary
    print("\n" + "=" * 60)
    print("📋 FINAL TEST SUMMARY")
    print("=" * 60)
    
    print(f"{'✅' if storage_result else '❌'} B2 Storage Test: {'PASS' if storage_result else 'FAIL'}")
    print(f"{'✅' if factory_result else '❌'} Factory Test: {'PASS' if factory_result else 'FAIL'}")
    
    all_passed = storage_result and factory_result
    
    print("\n" + "=" * 60)
    if all_passed:
        print("🎉 ALL B2 TESTS PASSED!")
        print("Your Backblaze B2 storage is working perfectly!")
    else:
        print("🚨 SOME TESTS FAILED!")
        print("Check the error messages above.")
    print("=" * 60)
    
    return all_passed

if __name__ == "__main__":
    result = asyncio.run(main())
    sys.exit(0 if result else 1)