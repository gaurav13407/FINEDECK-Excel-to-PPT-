"""
Backend Compatibility Test for User Service
Tests integration with all backend components
"""

import sys
import os
from datetime import datetime
from bson import ObjectId

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'backend', 'app'))

def test_imports():
    """Test all imports are compatible"""
    print("🔍 Testing imports compatibility...")
    
    try:
        # Test model imports
        from models.user import (
            UserCreate, UserInDB, SubscriptionPlan, SubscriptionStatus, 
            TemplateCategory, PyObjectId, has_credits_remaning, 
            deduct_credits, get_credit_usage_stats
        )
        print("✅ User models imported successfully")
        
        from models.file import FileType, ProcessingStatus, FileUpload
        print("✅ File models imported successfully")
        
        from models.conversion import PowerPointTemplate, ConversionSettings
        print("✅ Conversion models imported successfully")
        
        # Test core imports
        from core.security import get_password_hash, verify_password
        print("✅ Security core imported successfully")
        
        from core.database import get_collection, get_database
        print("✅ Database core imported successfully")
        
        # Test service imports
        from services.user_service import (
            create_user, get_user_by_email, get_user_by_id, 
            authenticate_user, update_user_profile, deduct_user_credits,
            upgrade_subscription, get_user_usage_stats
        )
        print("✅ User service imported successfully")
        
        return True
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False

def test_enum_compatibility():
    """Test enum compatibility across models"""
    print("\n🔍 Testing enum compatibility...")
    
    try:
        from models.user import SubscriptionPlan, SubscriptionStatus, TemplateCategory
        from models.file import FileType, ProcessingStatus
        from models.conversion import LayoutType, CharType, DataMappingType
        
        # Test SubscriptionPlan enum values
        expected_plans = {"FREE", "BASIC", "PRO", "ENTERPRISE"}
        actual_plans = {plan.name for plan in SubscriptionPlan}
        
        if expected_plans == actual_plans:
            print("✅ SubscriptionPlan enum has correct values")
        else:
            print(f"❌ SubscriptionPlan mismatch. Expected: {expected_plans}, Got: {actual_plans}")
            return False
        
        # Test TemplateCategory is shared correctly
        from models.file import TemplateCategory as FileTemplateCategory
        from models.conversion import TemplateCategory as ConversionTemplateCategory
        
        if TemplateCategory == FileTemplateCategory == ConversionTemplateCategory:
            print("✅ TemplateCategory shared correctly across models")
        else:
            print("❌ TemplateCategory not shared properly")
            return False
        
        # Test all enums have valid values
        test_enums = [
            (SubscriptionPlan, "SubscriptionPlan"),
            (SubscriptionStatus, "SubscriptionStatus"),
            (TemplateCategory, "TemplateCategory"),
            (FileType, "FileType"),
            (ProcessingStatus, "ProcessingStatus"),
            (LayoutType, "LayoutType"),
            (CharType, "CharType"),
            (DataMappingType, "DataMappingType")
        ]
        
        for enum_class, name in test_enums:
            values = [item.value for item in enum_class]
            if len(values) > 0:
                print(f"✅ {name} has {len(values)} valid values: {values}")
            else:
                print(f"❌ {name} has no values")
                return False
        
        return True
        
    except Exception as e:
        print(f"❌ Enum compatibility error: {e}")
        return False

def test_model_structure():
    """Test model structure compatibility"""
    print("\n🔍 Testing model structure compatibility...")
    
    try:
        from models.user import UserCreate, UserInDB, SubscriptionPlan
        
        # Test UserCreate structure
        test_user_data = {
            "name": "Test User",
            "email": "test@example.com", 
            "password": "testpassword123"
        }
        
        user_create = UserCreate(**test_user_data)
        print("✅ UserCreate model structure valid")
        
        # Test UserInDB structure  
        test_user_db = {
            "_id": ObjectId(),
            "name": "Test User",
            "email": "test@example.com",
            "password_hash": "hashed_password",
            "is_active": True,
            "subscription": {
                "plan": SubscriptionPlan.FREE,
                "status": "active",
                "start_date": datetime.utcnow(),
                "end_date": None,
                "credits_remaining": 1
            },
            "usage_stats": {
                "total_conversions": 0,
                "this_month_conversions": 0, 
                "total_credits_used": 0,
                "last_conversion_date": None
            },
            "created_at": datetime.utcnow(),
            "updated_at": datetime.utcnow(),
            "last_login": None
        }
        
        user_in_db = UserInDB(**test_user_db)
        print("✅ UserInDB model structure valid")
        
        # Test PyObjectId compatibility
        from models.user import PyObjectId
        test_id = PyObjectId()
        assert isinstance(test_id, ObjectId)
        print("✅ PyObjectId working correctly")
        
        return True
        
    except Exception as e:
        print(f"❌ Model structure error: {e}")
        return False

def test_service_functions():
    """Test service function signatures"""
    print("\n🔍 Testing service function signatures...")
    
    try:
        from services.user_service import (
            create_user, get_user_by_email, get_user_by_id,
            authenticate_user, update_user_profile, deduct_user_credits,
            upgrade_subscription, get_user_usage_stats
        )
        
        # Test function signatures exist
        functions = [
            ("create_user", create_user),
            ("get_user_by_email", get_user_by_email),
            ("get_user_by_id", get_user_by_id),
            ("authenticate_user", authenticate_user),
            ("update_user_profile", update_user_profile),
            ("deduct_user_credits", deduct_user_credits),
            ("upgrade_subscription", upgrade_subscription),
            ("get_user_usage_stats", get_user_usage_stats)
        ]
        
        for name, func in functions:
            if callable(func):
                print(f"✅ {name} function exists and callable")
            else:
                print(f"❌ {name} function not callable")
                return False
        
        return True
        
    except Exception as e:
        print(f"❌ Service function error: {e}")
        return False

def test_security_integration():
    """Test security integration"""
    print("\n🔍 Testing security integration...")
    
    try:
        from core.security import get_password_hash, verify_password
        
        # Test password hashing
        test_password = "testpassword123"
        hashed = get_password_hash(test_password)
        
        if isinstance(hashed, str) and len(hashed) > 10:
            print("✅ Password hashing works")
        else:
            print("❌ Password hashing failed")
            return False
        
        # Test password verification
        if verify_password(test_password, hashed):
            print("✅ Password verification works")
        else:
            print("❌ Password verification failed")
            return False
        
        # Test wrong password
        if not verify_password("wrongpassword", hashed):
            print("✅ Wrong password correctly rejected")
        else:
            print("❌ Wrong password incorrectly accepted")
            return False
        
        return True
        
    except Exception as e:
        print(f"❌ Security integration error: {e}")
        return False

def test_storage_integration():
    """Test storage system integration"""
    print("\n🔍 Testing storage integration...")
    
    try:
        from storage.b2_storage import get_storage_client
        
        # Test storage client factory
        local_storage = get_storage_client(use_b2=False)
        print("✅ Local storage client created")
        
        # Test that B2 storage class exists
        from storage.b2_storage import BackblazeB2Storage, LocalStorage
        print("✅ Storage classes imported successfully")
        
        return True
        
    except Exception as e:
        print(f"❌ Storage integration error: {e}")
        return False

def run_compatibility_tests():
    """Run all compatibility tests"""
    print("="*70)
    print("BACKEND COMPATIBILITY TEST SUITE")
    print("="*70)
    
    tests = [
        ("Import Compatibility", test_imports),
        ("Enum Compatibility", test_enum_compatibility),
        ("Model Structure", test_model_structure),
        ("Service Functions", test_service_functions),
        ("Security Integration", test_security_integration),
        ("Storage Integration", test_storage_integration)
    ]
    
    passed = 0
    total = len(tests)
    failed_tests = []
    
    for test_name, test_func in tests:
        print(f"\n{'='*20} {test_name} {'='*20}")
        try:
            if test_func():
                passed += 1
                print(f"✅ {test_name} PASSED")
            else:
                failed_tests.append(test_name)
                print(f"❌ {test_name} FAILED")
        except Exception as e:
            failed_tests.append(test_name)
            print(f"❌ {test_name} FAILED with exception: {e}")
    
    # Summary
    print("\n" + "="*70)
    print("COMPATIBILITY TEST SUMMARY")
    print("="*70)
    print(f"Total tests: {total}")
    print(f"Passed: {passed}")
    print(f"Failed: {len(failed_tests)}")
    print(f"Success rate: {(passed/total)*100:.1f}%")
    
    if failed_tests:
        print(f"\nFailed tests:")
        for test in failed_tests:
            print(f"  - {test}")
    else:
        print("\n🎉 All compatibility tests passed!")
        print("✅ Your backend is fully compatible and ready!")
    
    return len(failed_tests) == 0

if __name__ == "__main__":
    success = run_compatibility_tests()
    exit(0 if success else 1)