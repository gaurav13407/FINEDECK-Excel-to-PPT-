"""
File Service Compatibility Test
Tests file service integration with all backend components
"""

import sys
import os
import asyncio
from datetime import datetime, UTC
from bson import ObjectId

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'backend', 'app'))

async def test_file_service_imports():
    """Test file service imports"""
    print("🔍 Testing File Service Imports")
    print("="*50)
    
    try:
        # Test file service imports
        from services.file_service import (
            validate_file_upload, upload_file, created_conversion_job,
            get_user_files, get_file_download_url, delete_file
        )
        print("✅ File service functions imported successfully")
        
        # Test model imports used in file service
        from models.file import (
            FileUpload, FileRecord, FileResponse, ConversionJob,
            FileType, ProcessingStatus, TemplateCategory,
            validate_template_access, calculate_credits_needed, get_subscription_limits
        )
        print("✅ File models imported successfully")
        
        return True
        
    except Exception as e:
        print(f"❌ Import error: {e}")
        return False

async def test_function_signatures():
    """Test function signatures"""
    print("\n📝 Testing Function Signatures")
    print("="*50)
    
    try:
        from services.file_service import (
            validate_file_upload, upload_file, created_conversion_job,
            get_user_files, get_file_download_url, delete_file
        )
        
        import inspect
        
        functions = {
            'validate_file_upload': validate_file_upload,
            'upload_file': upload_file,
            'created_conversion_job': created_conversion_job,
            'get_user_files': get_user_files,
            'get_file_download_url': get_file_download_url,
            'delete_file': delete_file
        }
        
        for name, func in functions.items():
            sig = inspect.signature(func)
            print(f"✅ {name}: {sig}")
        
        return True
        
    except Exception as e:
        print(f"❌ Function signature error: {e}")
        return False

async def test_model_compatibility():
    """Test model compatibility with file service"""
    print("\n🔗 Testing Model Compatibility")
    print("="*50)
    
    try:
        from models.file import FileUpload, TemplateCategory, FileType
        
        # Test FileUpload creation
        file_upload = FileUpload(
            filename="test.xlsx",
            file_size=1024*1024,  # 1MB
            file_type=FileType.XLSX,
            template_category=TemplateCategory.BASIC,
            template_name="Test Template"
        )
        print("✅ FileUpload model creation successful")
        
        # Test enum values
        print(f"✅ FileType values: {[ft.value for ft in FileType]}")
        print(f"✅ TemplateCategory values: {[tc.value for tc in TemplateCategory]}")
        
        return True
        
    except Exception as e:
        print(f"❌ Model compatibility error: {e}")
        return False

async def test_utility_functions():
    """Test utility functions from file models"""
    print("\n🔧 Testing Utility Functions")
    print("="*50)
    
    try:
        from models.file import (
            validate_template_access, calculate_credits_needed, 
            get_subscription_limits, TemplateCategory
        )
        
        # Test template access validation
        basic_access = validate_template_access(TemplateCategory.BASIC, "basic")
        premium_access = validate_template_access(TemplateCategory.PREMIUM, "basic")
        print(f"✅ Template access - Basic for basic plan: {basic_access}")
        print(f"✅ Template access - Premium for basic plan: {premium_access}")
        
        # Test credit calculation
        basic_credits = calculate_credits_needed(TemplateCategory.BASIC)
        premium_credits = calculate_credits_needed(TemplateCategory.PREMIUM)
        print(f"✅ Credits needed - Basic: {basic_credits}, Premium: {premium_credits}")
        
        # Test subscription limits
        basic_limits = get_subscription_limits("basic")
        pro_limits = get_subscription_limits("pro")
        print(f"✅ Basic limits: {basic_limits}")
        print(f"✅ Pro limits: {pro_limits}")
        
        return True
        
    except Exception as e:
        print(f"❌ Utility function error: {e}")
        return False

async def test_datetime_usage():
    """Test datetime usage consistency"""
    print("\n⏰ Testing DateTime Usage")
    print("="*50)
    
    try:
        # Check for deprecated datetime.utcnow() usage
        file_path = os.path.join(os.path.dirname(__file__), '..', 'src', 'backend', 'app', 'services', 'file_service.py')
        
        with open(file_path, 'r') as f:
            content = f.read()
        
        # Check for deprecated patterns
        deprecated_patterns = ['datetime.utcnow()', 'datetime.utcnow(UTC)']
        issues = []
        
        for pattern in deprecated_patterns:
            if pattern in content:
                issues.append(f"Found deprecated datetime usage: {pattern}")
        
        if issues:
            print("⚠️ DateTime issues found:")
            for issue in issues:
                print(f"  - {issue}")
            return False
        else:
            print("✅ No deprecated datetime usage found")
            return True
        
    except Exception as e:
        print(f"❌ DateTime check error: {e}")
        return False

async def test_syntax_errors():
    """Test for syntax errors in file service"""
    print("\n🔍 Testing Syntax")
    print("="*50)
    
    try:
        # Try to compile the file service
        file_path = os.path.join(os.path.dirname(__file__), '..', 'src', 'backend', 'app', 'services', 'file_service.py')
        
        with open(file_path, 'r') as f:
            content = f.read()
        
        # Check for syntax issues
        syntax_issues = []
        
        # Check for missing return statement
        if "get_file_download_url" in content:
            lines = content.split('\n')
            in_download_function = False
            for i, line in enumerate(lines):
                if "async def get_file_download_url" in line:
                    in_download_function = True
                elif in_download_function and line.startswith("async def "):
                    in_download_function = False
                elif in_download_function and "return download_url" in line:
                    break
            else:
                if in_download_function:
                    syntax_issues.append("get_file_download_url function missing return statement")
        
        # Check for semicolon instead of colon
        if ")->Optional[str];" in content:
            syntax_issues.append("Found semicolon instead of colon in function signature")
        
        # Check for collection name typos
        if "conversions_jobs" in content:
            syntax_issues.append("Collection name should be 'conversion_jobs' not 'conversions_jobs'")
        
        if syntax_issues:
            print("❌ Syntax issues found:")
            for issue in syntax_issues:
                print(f"  - {issue}")
            return False
        else:
            print("✅ No major syntax issues found")
            return True
        
    except Exception as e:
        print(f"❌ Syntax check error: {e}")
        return False

async def test_integration_dependencies():
    """Test integration with other services"""
    print("\n🔗 Testing Integration Dependencies")
    print("="*50)
    
    try:
        # Test user service integration
        from services.user_service import get_user_by_id, deduct_user_credits
        print("✅ User service integration available")
        
        # Test storage integration
        from storage.b2_storage import get_storage_client
        print("✅ Storage service integration available")
        
        # Test database integration
        from core.database import get_collection
        print("✅ Database integration available")
        
        return True
        
    except Exception as e:
        print(f"❌ Integration dependency error: {e}")
        return False

async def run_file_service_tests():
    """Run all file service tests"""
    print("🚀 FILE SERVICE COMPATIBILITY TEST SUITE")
    print("="*70)
    
    tests = [
        ("Import Test", test_file_service_imports),
        ("Function Signatures", test_function_signatures),
        ("Model Compatibility", test_model_compatibility),
        ("Utility Functions", test_utility_functions),
        ("DateTime Usage", test_datetime_usage),
        ("Syntax Check", test_syntax_errors),
        ("Integration Dependencies", test_integration_dependencies)
    ]
    
    passed = 0
    total = len(tests)
    failed_tests = []
    
    for test_name, test_func in tests:
        print(f"\n{'='*15} {test_name} {'='*15}")
        try:
            if await test_func():
                passed += 1
                print(f"✅ {test_name} PASSED")
            else:
                failed_tests.append(test_name)
                print(f"❌ {test_name} FAILED")
        except Exception as e:
            failed_tests.append(test_name)
            print(f"❌ {test_name} FAILED with exception: {e}")
    
    # Summary
    print(f"\n{'='*70}")
    print("FILE SERVICE TEST SUMMARY")
    print("="*70)
    print(f"Total tests: {total}")
    print(f"Passed: {passed}")
    print(f"Failed: {len(failed_tests)}")
    print(f"Success rate: {(passed/total)*100:.1f}%")
    
    if failed_tests:
        print(f"\n❌ Failed tests:")
        for test in failed_tests:
            print(f"  - {test}")
        print(f"\n🔧 Issues to fix:")
        print(f"  1. Fix syntax errors (semicolon, missing returns)")
        print(f"  2. Update deprecated datetime usage")
        print(f"  3. Fix collection name typos")
    else:
        print("\n🎉 All file service tests passed!")
        print("✅ File service is compatible with backend!")
    
    return len(failed_tests) == 0

if __name__ == "__main__":
    success = asyncio.run(run_file_service_tests())
    exit(0 if success else 1)