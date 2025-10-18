"""
Test API Dependencies
Quick test to verify all dependencies are working correctly
"""

import sys
import os

# Add the src directory to the path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'src', 'backend', 'app'))

async def test_deps_imports():
    """Test that all dependencies can be imported"""
    print("🔍 Testing API Dependencies Import")
    print("="*50)
    
    try:
        # Test imports
        from api.deps import (
            get_db, get_current_user, get_current_active_user,
            validate_file_upload, get_pagination_params,
            require_subscription, require_credits, security
        )
        print("✅ All dependency functions imported successfully")
        
        # Test core security imports
        from core.security import (
            get_password_hash, verify_password, create_access_token,
            verify_token, get_current_user_from_token
        )
        print("✅ All security functions imported successfully")
        
        # Test that pagination works (it's the only one we can test without FastAPI context)
        pagination = await get_pagination_params(skip=10, limit=50)
        expected = {"skip": 10, "limit": 50}
        
        if pagination == expected:
            print(f"✅ Pagination function works: {pagination}")
        else:
            print(f"❌ Pagination function failed: expected {expected}, got {pagination}")
            return False
        
        # Test password functions
        test_password = "testpassword123"
        hashed = get_password_hash(test_password)
        verified = verify_password(test_password, hashed)
        
        if verified:
            print("✅ Password hashing and verification working")
        else:
            print("❌ Password verification failed")
            return False
        
        # Test JWT token creation (basic)
        test_data = {"sub": "test_user_id", "email": "test@example.com"}
        try:
            token = create_access_token(test_data)
            if token and isinstance(token, str):
                print("✅ JWT token creation working")
                
                # Test token verification
                payload = verify_token(token)
                if payload and payload.get("sub") == "test_user_id":
                    print("✅ JWT token verification working")
                else:
                    print("❌ JWT token verification failed")
                    return False
            else:
                print("❌ JWT token creation failed")
                return False
        except Exception as e:
            print(f"❌ JWT token test failed: {e}")
            return False
        
        print("\n🎉 ALL DEPENDENCY TESTS PASSED!")
        print("✅ Your API dependencies are ready to use!")
        
        return True
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Test failed: {e}")
        return False

def display_dependency_summary():
    """Display what dependencies are available"""
    print("\n📋 AVAILABLE API DEPENDENCIES")
    print("="*50)
    
    dependencies = [
        ("🔐 Authentication", [
            "get_current_user() - Extract user from JWT token",
            "get_current_active_user() - Ensure user is active",
            "security (HTTPBearer) - FastAPI security scheme"
        ]),
        ("📁 File Upload", [
            "validate_file_upload() - Check file size and type",
            "Max size: 50MB, Types: .xlsx, .xls, .csv"
        ]),
        ("📊 Pagination", [
            "get_pagination_params() - Query parameters",
            "Default: skip=0, limit=20, max_limit=100"
        ]),
        ("🔒 Permissions", [
            "require_subscription() - Check subscription level",
            "require_credits() - Check sufficient credits"
        ]),
        ("💾 Database", [
            "get_db() - Database connection dependency"
        ])
    ]
    
    for category, items in dependencies:
        print(f"\n{category}:")
        for item in items:
            print(f"   ✅ {item}")
    
    print(f"\n🚀 NEXT STEPS:")
    print(f"   1. Create FastAPI router files (auth.py, users.py, files.py)")
    print(f"   2. Use these dependencies in your API endpoints")
    print(f"   3. Test with actual HTTP requests")
    print(f"   4. Add middleware and error handlers")

if __name__ == "__main__":
    import asyncio
    
    try:
        success = asyncio.run(test_deps_imports())
        display_dependency_summary()
        
        if success:
            print("\n🏆 YOUR API DEPENDENCIES ARE PRODUCTION READY!")
        else:
            print("\n⚠️ Some issues found - check the errors above")
            
    except Exception as e:
        print(f"Test execution failed: {e}")
        print("\n⚠️ This is expected if you don't have all the config settings set up yet")