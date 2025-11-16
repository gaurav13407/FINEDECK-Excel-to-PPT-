# Test Dependencies Functions
from src.backend.app.core.dependencies import get_current_user, get_current_active_user, get_db
from src.backend.app.core.security import create_access_token
from fastapi.security import HTTPAuthorizationCredentials
from fastapi import HTTPException
import pytest

def test_get_db():
    """Test database connection dependency"""
    try:
        db = get_db()
        print("✅ Database dependency: WORKING")
        print(f"Database object: {type(db)}")
        return True
    except Exception as e:
        print(f"❌ Database dependency failed: {e}")
        return False

async def test_get_current_user_valid_token():
    """Test get_current_user with valid token"""
    try:
        # Create a valid token
        user_data = {"user_id": "test123", "email": "test@example.com"}
        token = create_access_token(user_data)
        
        # Create mock credentials
        class MockCredentials:
            def __init__(self, token):
                self.credentials = token
        
        credentials = MockCredentials(token)
        
        # Test the function
        result = await get_current_user(credentials)
        print("✅ Valid token test: WORKING")
        print(f"Returned user data: {result}")
        return True
        
    except Exception as e:
        print(f"❌ Valid token test failed: {e}")
        return False

async def test_get_current_user_invalid_token():
    """Test get_current_user with invalid token"""
    try:
        # Create mock credentials with invalid token
        class MockCredentials:
            def __init__(self, token):
                self.credentials = token
        
        credentials = MockCredentials("invalid_token_here")
        
        # Test the function - should raise HTTPException
        try:
            result = await get_current_user(credentials)
            print("❌ Invalid token test failed: Should have raised exception")
            return False
        except HTTPException as e:
            print("✅ Invalid token test: WORKING")
            print(f"Correctly raised HTTPException: {e.status_code} - {e.detail}")
            return True
        
    except Exception as e:
        print(f"❌ Invalid token test failed: {e}")
        return False

async def test_get_current_active_user():
    """Test get_current_active_user dependency chain"""
    try:
        # Mock user data
        mock_user = {"user_id": "test123", "email": "test@example.com"}
        
        # Test the function
        result = await get_current_active_user(mock_user)
        print("✅ Active user test: WORKING")
        print(f"Returned: {result}")
        return True
        
    except Exception as e:
        print(f"❌ Active user test failed: {e}")
        return False

# Run all tests
async def run_all_tests():
    print("🧪 Testing Dependencies...")
    print("=" * 50)
    
    # Test database
    db_test = test_get_db()
    
    # Test valid token
    valid_token_test = await test_get_current_user_valid_token()
    
    # Test invalid token
    invalid_token_test = await test_get_current_user_invalid_token()
    
    # Test active user
    active_user_test = await test_get_current_active_user()
    
    print("\n" + "=" * 50)
    print("📊 Test Results:")
    print(f"Database dependency: {'✅ PASS' if db_test else '❌ FAIL'}")
    print(f"Valid token handling: {'✅ PASS' if valid_token_test else '❌ FAIL'}")
    print(f"Invalid token handling: {'✅ PASS' if invalid_token_test else '❌ FAIL'}")
    print(f"Active user dependency: {'✅ PASS' if active_user_test else '❌ FAIL'}")
    
    if all([db_test, valid_token_test, invalid_token_test, active_user_test]):
        print("\n🎉 ALL DEPENDENCIES WORKING CORRECTLY!")
    else:
        print("\n⚠️ Some dependencies need fixes")

if __name__ == "__main__":
    import asyncio
    asyncio.run(run_all_tests())