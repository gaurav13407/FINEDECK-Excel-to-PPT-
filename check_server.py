"""
Quick server check - verify what's actually running
"""
import requests

BASE_URL = "http://localhost:8000"

print("=" * 80)
print("SERVER HEALTH CHECK")
print("=" * 80)

# Check 1: Root endpoint
print("\n1. Root endpoint (/):")
try:
    r = requests.get(f"{BASE_URL}/", timeout=5)
    print(f"   Status: {r.status_code}")
    if r.status_code == 200:
        print(f"   Response: {r.json()}")
        print("   ✅ Server is running!")
    else:
        print("   ❌ Unexpected status")
except requests.exceptions.ConnectionError:
    print("   ❌ Server not running!")
    print("\n   Start server with:")
    print("   cd src\\backend")
    print("   python run_server.py")
    exit(1)
except Exception as e:
    print(f"   ❌ Error: {e}")
    exit(1)

# Check 2: Health endpoint
print("\n2. Health endpoint (/health):")
try:
    r = requests.get(f"{BASE_URL}/health", timeout=5)
    print(f"   Status: {r.status_code}")
    if r.status_code == 200:
        print(f"   Response: {r.json()}")
        print("   ✅ Health check passed!")
except Exception as e:
    print(f"   ❌ Error: {e}")

# Check 3: Docs endpoint
print("\n3. API Docs (/docs):")
try:
    r = requests.get(f"{BASE_URL}/docs", timeout=5)
    print(f"   Status: {r.status_code}")
    if r.status_code == 200:
        print("   ✅ Docs accessible!")
    else:
        print(f"   ⚠️  Status {r.status_code}")
except Exception as e:
    print(f"   ❌ Error: {e}")

# Check 4: Session login endpoint
print("\n4. Session Login Endpoint (/api/v1/session/login):")
try:
    r = requests.post(
        f"{BASE_URL}/api/v1/session/login",
        json={"username": "admin", "password": "password"},
        timeout=5
    )
    print(f"   Status: {r.status_code}")
    if r.status_code == 200:
        print("   ✅ Session endpoint registered!")
        print(f"   Response: {r.json()}")
    elif r.status_code == 404:
        print("   ❌ Endpoint not found - router not registered")
    elif r.status_code == 401:
        print("   ⚠️  Endpoint exists but credentials invalid")
    else:
        print(f"   Response: {r.text[:200]}")
except Exception as e:
    print(f"   ❌ Error: {e}")

# Check 5: Session protected endpoint
print("\n5. Session Protected Endpoint (/api/v1/session/protected):")
try:
    r = requests.get(f"{BASE_URL}/api/v1/session/protected", timeout=5)
    print(f"   Status: {r.status_code}")
    if r.status_code == 401:
        print("   ✅ Protected endpoint working (requires auth)")
        print(f"   Message: {r.json()}")
    elif r.status_code == 404:
        print("   ❌ Endpoint not found")
    else:
        print(f"   Response: {r.text[:200]}")
except Exception as e:
    print(f"   ❌ Error: {e}")

print("\n" + "=" * 80)
print("Check server console for:")
print("  • 'Connected to Redis successfully!'")
print("  • 'Session middleware enabled'")
print("=" * 80)
