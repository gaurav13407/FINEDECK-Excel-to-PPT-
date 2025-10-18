#!/usr/bin/env python3
"""
Test Auth and Users API Endpoints
Tests both authentication and user management functionality
"""
import sys
import os
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

def test_auth_endpoints_imports():
    """Test that auth endpoints can be imported"""
    print("🔍 Testing Auth Endpoints Imports")
    print("=" * 50)
    
    try:
        from src.backend.app.api.v1.endpoints.auth import router as auth_router
        print("✅ Auth router imported successfully")
        
        # Check router configuration
        print(f"📍 Auth prefix: {auth_router.prefix}")
        print(f"🏷️  Auth tags: {auth_router.tags}")
        
        # Count routes
        route_count = len(auth_router.routes)
        print(f"🛤️  Auth routes count: {route_count}")
        
        # List route paths
        print("📋 Auth Routes:")
        for route in auth_router.routes:
            if hasattr(route, 'path') and hasattr(route, 'methods'):
                methods = ', '.join(route.methods)
                print(f"   {route.path} [{methods}]")
        
        return True
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False


def test_users_endpoints_imports():
    """Test that users endpoints can be imported"""
    print("\n🔍 Testing Users Endpoints Imports")
    print("=" * 50)
    
    try:
        from src.backend.app.api.v1.endpoints.users import router as users_router
        print("✅ Users router imported successfully")
        
        # Check router configuration
        print(f"📍 Users prefix: {users_router.prefix}")
        print(f"🏷️  Users tags: {users_router.tags}")
        
        # Count routes
        route_count = len(users_router.routes)
        print(f"🛤️  Users routes count: {route_count}")
        
        # List route paths
        print("📋 Users Routes:")
        for route in users_router.routes:
            if hasattr(route, 'path') and hasattr(route, 'methods'):
                methods = ', '.join(route.methods)
                print(f"   {route.path} [{methods}]")
        
        return True
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False


def test_dependencies_compatibility():
    """Test that dependencies work with endpoints"""
    print("\n🔍 Testing Dependencies Compatibility")
    print("=" * 50)
    
    try:
        # Test core dependencies
        from src.backend.app.api.deps import (
            get_current_user, get_current_active_user, 
            validate_file_upload, require_credits
        )
        print("✅ Core dependencies imported successfully")
        
        # Test models
        from src.backend.app.models.user import (
            UserCreate, UserUpdate, UserResponse, Token
        )
        print("✅ User models imported successfully")
        
        # Test services
        from src.backend.app.services.user_service import (
            create_user, get_user_by_email, authenticate_user,
            update_user_profile, deduct_user_credits
        )
        print("✅ User services imported successfully")
        
        # Test security
        from src.backend.app.core.security import (
            create_access_token, verify_password
        )
        print("✅ Security functions imported successfully")
        
        return True
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False


def test_model_compatibility():
    """Test that models are properly structured"""
    print("\n🔍 Testing Model Compatibility")
    print("=" * 50)
    
    try:
        from src.backend.app.models.user import UserUpdate, UserCreate, UserResponse
        
        # Test UserCreate model
        user_create_data = {
            "name": "Test User",
            "email": "test@example.com",
            "password": "password123"
        }
        user_create = UserCreate(**user_create_data)
        print("✅ UserCreate model works correctly")
        print(f"   Name: {user_create.name}")
        print(f"   Email: {user_create.email}")
        
        # Test UserUpdate model
        user_update_data = {
            "name": "Updated Name",
            "is_active": True
        }
        user_update = UserUpdate(**user_update_data)
        print("✅ UserUpdate model works correctly")
        print(f"   Updated fields: {user_update.dict(exclude_unset=True)}")
        
        return True
    except Exception as e:
        print(f"❌ Model error: {e}")
        return False


def test_endpoint_structure():
    """Test endpoint structure and dependencies"""
    print("\n🔍 Testing Endpoint Structure")
    print("=" * 50)
    
    try:
        # Check if we can create a simple FastAPI app with our routers
        from fastapi import FastAPI
        from src.backend.app.api.v1.endpoints.auth import router as auth_router
        from src.backend.app.api.v1.endpoints.users import router as users_router
        
        app = FastAPI()
        app.include_router(auth_router, prefix="/api/v1")
        app.include_router(users_router, prefix="/api/v1")
        
        print("✅ FastAPI app created successfully with routers")
        
        # Count total routes
        total_routes = len(app.routes)
        print(f"📊 Total routes in app: {total_routes}")
        
        # List all routes
        print("📋 All API Routes:")
        for route in app.routes:
            if hasattr(route, 'path') and hasattr(route, 'methods'):
                methods = ', '.join(route.methods)
                print(f"   {route.path} [{methods}]")
        
        return True
    except Exception as e:
        print(f"❌ Structure error: {e}")
        return False


def show_api_usage_examples():
    """Show example API requests"""
    print("\n🔍 API Usage Examples")
    print("=" * 50)
    
    print("🔐 AUTH ENDPOINTS:")
    print("   POST /api/v1/auth/signup")
    print('   Body: {"name": "John Doe", "email": "john@example.com", "password": "password123"}')
    print()
    print("   POST /api/v1/auth/login")
    print('   Body: {"username": "john@example.com", "password": "password123"}')
    print()
    print("   GET /api/v1/auth/me")
    print("   Headers: Authorization: Bearer <your-jwt-token>")
    print()
    
    print("👤 USER ENDPOINTS:")
    print("   GET /api/v1/users/me")
    print("   Headers: Authorization: Bearer <your-jwt-token>")
    print()
    print("   PUT /api/v1/users/me")
    print("   Headers: Authorization: Bearer <your-jwt-token>")
    print('   Body: {"name": "Updated Name", "email": "newemail@example.com"}')
    print()
    print("   GET /api/v1/users/stats")
    print("   Headers: Authorization: Bearer <your-jwt-token>")
    print()
    print("   POST /api/v1/users/credits/deduct?credits_amount=1")
    print("   Headers: Authorization: Bearer <your-jwt-token>")


def main():
    """Run all tests"""
    print("🚀 FINDECK AUTH & USERS API TESTS")
    print("=" * 80)
    
    tests = [
        ("Auth Endpoints Import", test_auth_endpoints_imports),
        ("Users Endpoints Import", test_users_endpoints_imports),
        ("Dependencies Compatibility", test_dependencies_compatibility),
        ("Model Compatibility", test_model_compatibility),
        ("Endpoint Structure", test_endpoint_structure),
    ]
    
    results = []
    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ Test '{test_name}' failed with error: {e}")
            results.append((test_name, False))
    
    # Show usage examples
    show_api_usage_examples()
    
    # Summary
    print("\n📊 TEST SUMMARY")
    print("=" * 60)
    passed = sum(1 for _, result in results if result)
    total = len(results)
    
    for test_name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"   {test_name}: {status}")
    
    print(f"\n🎯 Overall: {passed}/{total} tests passed")
    
    if passed == total:
        print("\n🎉 ALL TESTS PASSED!")
        print("\n✅ Your Auth and Users endpoints are ready!")
        print("\n🚀 Next steps:")
        print("   1. Set up your main FastAPI app")
        print("   2. Include these routers")
        print("   3. Connect to MongoDB")
        print("   4. Test with real HTTP requests")
        print("\n📖 Start server with:")
        print("   python -m uvicorn main:app --reload")
    else:
        print(f"\n⚠️  {total - passed} tests failed. Check the errors above.")


if __name__ == "__main__":
    main()