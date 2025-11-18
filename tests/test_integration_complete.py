#!/usr/bin/env python3
"""
Comprehensive test script for FinDeck Redis Session Integration.
Tests:
1. Server startup
2. Redis connection
3. Session creation and management
4. Middleware protection
5. Authentication flow
6. Protected endpoints
"""

import os
import sys
import time
import requests
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Test configuration
BASE_URL = "http://localhost:8000"
TEST_USERNAME = "test_user"
TEST_PASSWORD = "test_password"


class Colors:
    """ANSI color codes for terminal output."""
    GREEN = '\033[92m'
    RED = '\033[91m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    RESET = '\033[0m'
    BOLD = '\033[1m'


def print_header(text):
    """Print formatted section header."""
    print(f"\n{Colors.BOLD}{Colors.BLUE}{'=' * 80}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{text}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.BLUE}{'=' * 80}{Colors.RESET}\n")


def print_success(text):
    """Print success message."""
    print(f"{Colors.GREEN}✅ {text}{Colors.RESET}")


def print_error(text):
    """Print error message."""
    print(f"{Colors.RED}❌ {text}{Colors.RESET}")


def print_warning(text):
    """Print warning message."""
    print(f"{Colors.YELLOW}⚠️  {text}{Colors.RESET}")


def print_info(text):
    """Print info message."""
    print(f"{Colors.BLUE}ℹ️  {text}{Colors.RESET}")


def test_server_connection():
    """Test if server is running and accessible."""
    print_header("Test 1: Server Connection")
    
    try:
        response = requests.get(f"{BASE_URL}/", timeout=5)
        
        if response.status_code == 200:
            print_success(f"Server is running at {BASE_URL}")
            print_info(f"Response: {response.json()}")
            return True
        else:
            print_error(f"Server returned status code: {response.status_code}")
            return False
    except requests.ConnectionError:
        print_error(f"Cannot connect to server at {BASE_URL}")
        print_warning("Please start the server with: python src/backend/run_server.py")
        return False
    except Exception as e:
        print_error(f"Unexpected error: {e}")
        return False


def test_api_docs():
    """Test if API documentation is accessible."""
    print_header("Test 2: API Documentation")
    
    try:
        response = requests.get(f"{BASE_URL}/docs", timeout=5)
        
        if response.status_code == 200:
            print_success("API docs accessible at /docs")
            print_info(f"Documentation size: {len(response.content)} bytes")
            return True
        else:
            print_error(f"API docs returned status code: {response.status_code}")
            return False
    except Exception as e:
        print_error(f"Error accessing API docs: {e}")
        return False


def test_protected_route_without_auth():
    """Test that protected routes require authentication."""
    print_header("Test 3: Protected Routes (Without Auth)")
    
    # Try accessing a protected route without session
    protected_endpoints = [
        "/api/v1/users/me",
        "/api/v1/files",
        "/api/v1/session/protected",
    ]
    
    all_protected = True
    for endpoint in protected_endpoints:
        try:
            response = requests.get(f"{BASE_URL}{endpoint}", timeout=5)
            
            if response.status_code == 401:
                print_success(f"{endpoint} is protected (401 Unauthorized)")
            elif response.status_code == 404:
                print_info(f"{endpoint} not found (404) - endpoint may not exist yet")
            else:
                print_warning(f"{endpoint} returned {response.status_code} - expected 401")
                all_protected = False
        except Exception as e:
            print_error(f"Error testing {endpoint}: {e}")
            all_protected = False
    
    return all_protected


def test_login():
    """Test login endpoint and session creation."""
    print_header("Test 4: Login & Session Creation")
    
    try:
        # Test with session example endpoint
        response = requests.post(
            f"{BASE_URL}/api/v1/session/login",
            json={
                "username": "admin",
                "password": "password"
            },
            timeout=5
        )
        
        if response.status_code == 200:
            data = response.json()
            print_success("Login successful")
            print_info(f"Session ID: {data.get('session_id', 'N/A')[:20]}...")
            print_info(f"User: {data.get('user', {})}")
            
            # Check if cookie was set
            if 'session_id' in response.cookies:
                print_success("Session cookie set successfully")
                return True, response.cookies
            else:
                print_warning("Login succeeded but no session cookie set")
                return True, None
        else:
            print_error(f"Login failed with status code: {response.status_code}")
            print_error(f"Response: {response.text}")
            return False, None
    except requests.exceptions.ConnectionError:
        print_error("Cannot connect to server")
        return False, None
    except Exception as e:
        print_error(f"Login error: {e}")
        return False, None


def test_protected_route_with_auth(cookies):
    """Test accessing protected route with valid session."""
    print_header("Test 5: Protected Routes (With Auth)")
    
    if not cookies:
        print_warning("No cookies available - skipping authenticated tests")
        return False
    
    try:
        response = requests.get(
            f"{BASE_URL}/api/v1/session/protected",
            cookies=cookies,
            timeout=5
        )
        
        if response.status_code == 200:
            data = response.json()
            print_success("Successfully accessed protected route")
            print_info(f"Response: {data}")
            return True
        else:
            print_error(f"Protected route returned status code: {response.status_code}")
            print_error(f"Response: {response.text}")
            return False
    except Exception as e:
        print_error(f"Error accessing protected route: {e}")
        return False


def test_get_current_user(cookies):
    """Test getting current user from session."""
    print_header("Test 6: Get Current User")
    
    if not cookies:
        print_warning("No cookies available - skipping test")
        return False
    
    try:
        response = requests.get(
            f"{BASE_URL}/api/v1/session/me",
            cookies=cookies,
            timeout=5
        )
        
        if response.status_code == 200:
            user_data = response.json()
            print_success("Successfully retrieved current user")
            print_info(f"User ID: {user_data.get('user_id')}")
            print_info(f"Username: {user_data.get('username')}")
            print_info(f"Email: {user_data.get('email')}")
            return True
        else:
            print_error(f"Get current user failed: {response.status_code}")
            print_error(f"Response: {response.text}")
            return False
    except Exception as e:
        print_error(f"Error getting current user: {e}")
        return False


def test_session_refresh(cookies):
    """Test session refresh functionality."""
    print_header("Test 7: Session Refresh")
    
    if not cookies:
        print_warning("No cookies available - skipping test")
        return False
    
    try:
        response = requests.post(
            f"{BASE_URL}/api/v1/session/refresh",
            cookies=cookies,
            timeout=5
        )
        
        if response.status_code == 200:
            data = response.json()
            print_success("Session refreshed successfully")
            print_info(f"Response: {data}")
            return True
        else:
            print_error(f"Session refresh failed: {response.status_code}")
            return False
    except Exception as e:
        print_error(f"Error refreshing session: {e}")
        return False


def test_logout(cookies):
    """Test logout and session destruction."""
    print_header("Test 8: Logout & Session Destruction")
    
    if not cookies:
        print_warning("No cookies available - skipping test")
        return False
    
    try:
        response = requests.post(
            f"{BASE_URL}/api/v1/session/logout",
            cookies=cookies,
            timeout=5
        )
        
        if response.status_code == 200:
            data = response.json()
            print_success("Logout successful")
            print_info(f"Response: {data}")
            
            # Verify session is destroyed by trying to access protected route
            verify_response = requests.get(
                f"{BASE_URL}/api/v1/session/protected",
                cookies=cookies,
                timeout=5
            )
            
            if verify_response.status_code == 401:
                print_success("Session properly destroyed (401 on protected route)")
                return True
            else:
                print_warning(f"Session may still be valid (got {verify_response.status_code})")
                return False
        else:
            print_error(f"Logout failed: {response.status_code}")
            return False
    except Exception as e:
        print_error(f"Error during logout: {e}")
        return False


def test_middleware_configuration():
    """Test that middleware is properly configured."""
    print_header("Test 9: Middleware Configuration")
    
    try:
        # Test that public routes are accessible
        public_routes = [
            ("/", "Homepage"),
            ("/docs", "API Documentation"),
        ]
        
        all_accessible = True
        for route, name in public_routes:
            try:
                response = requests.get(f"{BASE_URL}{route}", timeout=5)
                if response.status_code == 200:
                    print_success(f"{name} ({route}) is publicly accessible")
                else:
                    print_warning(f"{name} ({route}) returned {response.status_code}")
                    all_accessible = False
            except Exception as e:
                print_error(f"Error accessing {route}: {e}")
                all_accessible = False
        
        return all_accessible
    except Exception as e:
        print_error(f"Error testing middleware: {e}")
        return False


def test_redis_backend():
    """Test that Redis is properly configured and connected."""
    print_header("Test 10: Redis Backend Connection")
    
    import redis
    
    redis_url = os.getenv("UPSTASH_REDIS_URL")
    
    if not redis_url:
        print_error("UPSTASH_REDIS_URL not found in environment")
        return False
    
    try:
        client = redis.from_url(redis_url)
        
        # Test ping
        if client.ping():
            print_success("Redis PING successful")
        else:
            print_error("Redis PING failed")
            return False
        
        # Test set/get with session-like key
        test_key = "session:test_integration_123"
        test_value = '{"user_id": "test", "username": "integration_test"}'
        
        client.set(test_key, test_value, ex=60)
        retrieved = client.get(test_key)
        
        if retrieved and retrieved.decode() == test_value:
            print_success("Redis SET/GET working correctly")
            client.delete(test_key)
            print_success("Redis session storage verified")
            return True
        else:
            print_error("Redis SET/GET failed")
            return False
    except Exception as e:
        print_error(f"Redis connection error: {e}")
        return False


def main():
    """Run all tests."""
    print(f"\n{Colors.BOLD}{'=' * 80}")
    print(f"FinDeck Redis Session Integration - Comprehensive Test Suite")
    print(f"{'=' * 80}{Colors.RESET}\n")
    
    print_info(f"Testing server at: {BASE_URL}")
    print_info(f"Test started at: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    results = {}
    cookies = None
    
    # Test 1: Server Connection
    results['Server Connection'] = test_server_connection()
    
    if not results['Server Connection']:
        print_error("\n❌ Server is not running. Please start the server first:")
        print_info("   python src/backend/run_server.py")
        print_error("\nAborting remaining tests.\n")
        sys.exit(1)
    
    # Test 2: API Documentation
    results['API Documentation'] = test_api_docs()
    
    # Test 3: Protected Routes Without Auth
    results['Protected Routes (No Auth)'] = test_protected_route_without_auth()
    
    # Test 4: Login
    login_success, cookies = test_login()
    results['Login & Session Creation'] = login_success
    
    # Test 5: Protected Routes With Auth
    if cookies:
        results['Protected Routes (With Auth)'] = test_protected_route_with_auth(cookies)
        results['Get Current User'] = test_get_current_user(cookies)
        results['Session Refresh'] = test_session_refresh(cookies)
        results['Logout'] = test_logout(cookies)
    else:
        print_warning("\nSkipping authenticated tests - no session cookies available")
    
    # Test 9: Middleware Configuration
    results['Middleware Configuration'] = test_middleware_configuration()
    
    # Test 10: Redis Backend
    results['Redis Backend'] = test_redis_backend()
    
    # Print summary
    print_header("Test Results Summary")
    
    passed = 0
    failed = 0
    
    for test_name, result in results.items():
        status = f"{Colors.GREEN}✅ PASSED{Colors.RESET}" if result else f"{Colors.RED}❌ FAILED{Colors.RESET}"
        print(f"{test_name:.<50} {status}")
        
        if result:
            passed += 1
        else:
            failed += 1
    
    print(f"\n{Colors.BOLD}{'=' * 80}{Colors.RESET}")
    print(f"{Colors.BOLD}Total Tests: {passed + failed} | Passed: {Colors.GREEN}{passed}{Colors.RESET} | Failed: {Colors.RED}{failed}{Colors.RESET}{Colors.BOLD}{Colors.RESET}")
    print(f"{Colors.BOLD}{'=' * 80}{Colors.RESET}\n")
    
    if failed == 0:
        print_success("🎉 All tests passed! Redis session integration is working correctly.\n")
        sys.exit(0)
    else:
        print_error(f"⚠️  {failed} test(s) failed. Please review the output above.\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
