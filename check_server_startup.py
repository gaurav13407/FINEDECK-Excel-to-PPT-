#!/usr/bin/env python3
"""
Check if server is running and properly configured with Redis.
"""
import requests
import json

BASE_URL = "http://localhost:8000"

print("=" * 80)
print("SERVER STARTUP CHECK")
print("=" * 80)

# Test 1: Check if server is running
print("\n1. Checking if server is running...")
try:
    response = requests.get(f"{BASE_URL}/docs", timeout=5)
    if response.status_code == 200:
        print("   ✅ Server is running")
    else:
        print(f"   ❌ Unexpected status: {response.status_code}")
except requests.exceptions.ConnectionError:
    print("   ❌ FAIL - Server not running!")
    print("   → Start server with: cd src\\backend && python run_server.py")
    exit(1)
except Exception as e:
    print(f"   ❌ FAIL - {e}")
    exit(1)

# Test 2: Check if session endpoints exist
print("\n2. Checking session endpoints...")
endpoints_to_check = [
    "/api/v1/session/login",
    "/api/v1/session/logout",
    "/api/v1/session/protected",
    "/api/v1/session/me",
]

all_exist = True
for endpoint in endpoints_to_check:
    try:
        # Try POST for login, GET for others
        if "login" in endpoint:
            response = requests.post(
                f"{BASE_URL}{endpoint}",
                json={"username": "test", "password": "test"}
            )
        else:
            response = requests.get(f"{BASE_URL}{endpoint}")
        
        if response.status_code in [401, 403, 422]:
            # These are good - endpoint exists but we're not authorized
            print(f"   ✅ {endpoint} - exists (needs auth)")
        elif response.status_code == 503:
            print(f"   ⚠️  {endpoint} - exists but Redis not configured")
        elif response.status_code == 404:
            print(f"   ❌ {endpoint} - NOT FOUND")
            all_exist = False
        else:
            print(f"   ✅ {endpoint} - exists (status: {response.status_code})")
    except Exception as e:
        print(f"   ❌ {endpoint} - Error: {e}")
        all_exist = False

# Test 3: Check OpenAPI spec for registered routes
print("\n3. Checking OpenAPI spec for session routes...")
try:
    response = requests.get(f"{BASE_URL}/openapi.json")
    if response.status_code == 200:
        openapi = response.json()
        paths = openapi.get("paths", {})
        
        session_routes = [p for p in paths.keys() if "/session/" in p]
        
        if session_routes:
            print(f"   ✅ Found {len(session_routes)} session routes:")
            for route in session_routes:
                print(f"      - {route}")
        else:
            print("   ❌ No session routes found in OpenAPI spec!")
            print("   → This means auth_example.py router is NOT registered")
            print("   → Check api/v1/api.py to ensure auth_example is included")
    else:
        print(f"   ❌ Failed to get OpenAPI spec: {response.status_code}")
except Exception as e:
    print(f"   ❌ Error: {e}")

print("\n" + "=" * 80)
if all_exist:
    print("✅ SERVER CONFIGURED CORRECTLY")
    print("\nNext step: Test the session flow with:")
    print("   python test_middleware_manual.py")
else:
    print("❌ SERVER CONFIGURATION ISSUES FOUND")
    print("\nTroubleshooting steps:")
    print("   1. Make sure you started server with: cd src\\backend && python run_server.py")
    print("   2. Check that auth_example.py router is registered in api/v1/api.py")
    print("   3. Restart the server after any code changes")
print("=" * 80)
