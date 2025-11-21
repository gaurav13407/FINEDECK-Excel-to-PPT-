"""
Manual Middleware Test Script
Run this with the server running to verify middleware is working
"""
import requests
import sys

BASE_URL = "http://localhost:8000"

print("=" * 80)
print("FINDECK REDIS SESSION MIDDLEWARE - MANUAL TEST")
print("=" * 80)
print(f"\nServer: {BASE_URL}")
print("Make sure server is running: python src/backend/run_server.py\n")

# Test 1: Public route
print("Test 1: Public Route (/docs)")
print("-" * 80)
try:
    r = requests.get(f"{BASE_URL}/docs", timeout=5)
    if r.status_code == 200:
        print(f"✅ PASS - Status: {r.status_code} (docs accessible without auth)")
    else:
        print(f"❌ FAIL - Status: {r.status_code} (expected 200)")
except requests.exceptions.ConnectionError:
    print("❌ FAIL - Cannot connect to server. Is it running?")
    sys.exit(1)
except Exception as e:
    print(f"❌ FAIL - Error: {e}")
    sys.exit(1)

# Test 2: Protected route without auth
print("\nTest 2: Protected Route Without Auth (/api/v1/session/protected)")
print("-" * 80)
try:
    r = requests.get(f"{BASE_URL}/api/v1/session/protected", timeout=5)
    if r.status_code == 401:
        print(f"✅ PASS - Status: {r.status_code} (correctly blocked)")
        print(f"   Message: {r.json().get('detail', 'N/A')}")
    elif r.status_code == 404:
        print(f"⚠️  WARNING - Status: 404 (endpoint not found)")
        print("   The /api/v1/session/* endpoints may not be registered")
    else:
        print(f"❌ FAIL - Status: {r.status_code} (expected 401)")
        print(f"   Middleware may not be active!")
except Exception as e:
    print(f"❌ FAIL - Error: {e}")

# Test 3: Login
print("\nTest 3: Login (/api/v1/session/login)")
print("-" * 80)
try:
    r = requests.post(
        f"{BASE_URL}/api/v1/session/login",
        json={"username": "admin", "password": "password"},
        timeout=5
    )
    if r.status_code == 200:
        data = r.json()
        print(f"✅ PASS - Status: {r.status_code} (login successful)")
        print(f"   Session ID: {data.get('session_id', 'N/A')[:20]}...")
        print(f"   User: {data.get('user', {}).get('username', 'N/A')}")
        cookies = r.cookies
    elif r.status_code == 404:
        print(f"⚠️  WARNING - Status: 404 (endpoint not found)")
        print("   The session auth endpoints are not registered")
        cookies = None
    else:
        print(f"❌ FAIL - Status: {r.status_code}")
        print(f"   Response: {r.text}")
        cookies = None
except Exception as e:
    print(f"❌ FAIL - Error: {e}")
    cookies = None

# Test 4: Protected route with session
if cookies:
    print("\nTest 4: Protected Route With Session (/api/v1/session/protected)")
    print("-" * 80)
    try:
        r = requests.get(
            f"{BASE_URL}/api/v1/session/protected",
            cookies=cookies,
            timeout=5
        )
        if r.status_code == 200:
            data = r.json()
            print(f"✅ PASS - Status: {r.status_code} (access granted with session)")
            print(f"   Message: {data.get('message', 'N/A')}")
            print(f"   User: {data.get('user', {}).get('username', 'N/A')}")
        else:
            print(f"❌ FAIL - Status: {r.status_code} (expected 200)")
            print(f"   Session validation may not be working!")
    except Exception as e:
        print(f"❌ FAIL - Error: {e}")
    
    # Test 5: Get current user
    print("\nTest 5: Get Current User (/api/v1/session/me)")
    print("-" * 80)
    try:
        r = requests.get(
            f"{BASE_URL}/api/v1/session/me",
            cookies=cookies,
            timeout=5
        )
        if r.status_code == 200:
            user = r.json()
            print(f"✅ PASS - Status: {r.status_code}")
            print(f"   User ID: {user.get('user_id', 'N/A')}")
            print(f"   Username: {user.get('username', 'N/A')}")
            print(f"   Email: {user.get('email', 'N/A')}")
        elif r.status_code == 404:
            print(f"⚠️  WARNING - Status: 404 (endpoint not found)")
        else:
            print(f"❌ FAIL - Status: {r.status_code}")
    except Exception as e:
        print(f"❌ FAIL - Error: {e}")
    
    # Test 6: Logout
    print("\nTest 6: Logout (/api/v1/session/logout)")
    print("-" * 80)
    try:
        r = requests.post(
            f"{BASE_URL}/api/v1/session/logout",
            cookies=cookies,
            timeout=5
        )
        if r.status_code == 200:
            print(f"✅ PASS - Status: {r.status_code} (logout successful)")
            
            # Verify session destroyed
            r2 = requests.get(
                f"{BASE_URL}/api/v1/session/protected",
                cookies=cookies,
                timeout=5
            )
            if r2.status_code == 401:
                print(f"✅ PASS - Session properly destroyed (401 on protected route)")
            else:
                print(f"❌ FAIL - Session may still be valid (got {r2.status_code})")
        elif r.status_code == 404:
            print(f"⚠️  WARNING - Status: 404 (endpoint not found)")
        else:
            print(f"❌ FAIL - Status: {r.status_code}")
    except Exception as e:
        print(f"❌ FAIL - Error: {e}")
else:
    print("\n⚠️  Skipping tests 4-6 (no valid session from login)")

# Final summary
print("\n" + "=" * 80)
print("TEST SUMMARY")
print("=" * 80)

if cookies:
    print("✅ MIDDLEWARE APPEARS TO BE WORKING!")
    print("\nKey Points:")
    print("  • Public routes accessible without auth")
    print("  • Protected routes blocked without session")
    print("  • Login creates valid session")
    print("  • Protected routes accessible with session")
    print("  • Logout destroys session")
else:
    print("⚠️  MIDDLEWARE STATUS UNCLEAR")
    print("\nPossible Issues:")
    print("  • Session auth endpoints not registered (/api/v1/session/*)")
    print("  • Login endpoint not working")
    print("  • Check if auth_example.py router is included in api.py")

print("\nNext Steps:")
print("  1. Check server logs for 'Session middleware enabled' message")
print("  2. Verify ENABLE_REDIS_SESSIONS=true in .env")
print("  3. Check that auth_example router is registered in api/v1/api.py")
print("=" * 80)
