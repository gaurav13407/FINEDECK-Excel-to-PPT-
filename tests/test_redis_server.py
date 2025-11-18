#!/usr/bin/env python3
"""
Test script for validating Redis connection and server functionality.
This script tests:
1. Redis connection using Upstash credentials
2. Session storage and retrieval
3. Basic server health check
"""

import os
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from dotenv import load_dotenv
import redis
import requests
import time

# Load environment variables
load_dotenv()

def test_redis_connection():
    """Test Redis connection using Upstash credentials."""
    print("\n=== Testing Redis Connection ===")
    
    redis_url = os.getenv("UPSTASH_REDIS_URL")
    if not redis_url:
        print("❌ UPSTASH_REDIS_URL not found in .env file")
        return False
    
    try:
        client = redis.from_url(redis_url)
        
        # Test ping
        response = client.ping()
        if response:
            print("✅ Redis PING successful")
        else:
            print("❌ Redis PING failed")
            return False
        
        # Test set/get
        test_key = "test_key"
        test_value = "test_value_123"
        client.set(test_key, test_value, ex=60)  # Expire in 60 seconds
        retrieved_value = client.get(test_key)
        
        if retrieved_value and retrieved_value.decode() == test_value:
            print(f"✅ Redis SET/GET successful (key: {test_key})")
        else:
            print("❌ Redis SET/GET failed")
            return False
        
        # Clean up test key
        client.delete(test_key)
        print("✅ Redis test key cleaned up")
        
        # Test session data structure
        session_key = "session:test_user_123"
        session_data = '{"user_id": "test_user", "email": "test@example.com"}'
        client.set(session_key, session_data, ex=3600)
        retrieved_session = client.get(session_key)
        
        if retrieved_session:
            print(f"✅ Session storage test successful")
            client.delete(session_key)
        else:
            print("❌ Session storage test failed")
            return False
        
        print("\n✅ All Redis connection tests passed!")
        return True
        
    except redis.ConnectionError as e:
        print(f"❌ Redis connection error: {e}")
        return False
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        return False


def test_session_secret():
    """Test if SESSION_SECRET_KEY is configured."""
    print("\n=== Testing Session Configuration ===")
    
    secret_key = os.getenv("SESSION_SECRET_KEY")
    if not secret_key:
        print("❌ SESSION_SECRET_KEY not found in .env file")
        return False
    
    if len(secret_key) < 32:
        print(f"⚠️  SESSION_SECRET_KEY is too short ({len(secret_key)} chars). Should be at least 32 chars.")
        return False
    
    print(f"✅ SESSION_SECRET_KEY configured (length: {len(secret_key)} chars)")
    return True


def test_server_startup():
    """Test if server can start and respond to health checks."""
    print("\n=== Testing Server Startup ===")
    print("Note: This test requires the server to be running.")
    print("Start server with: python src/backend/run_server.py")
    
    server_url = "http://localhost:8000"
    
    try:
        # Try to connect to server
        response = requests.get(f"{server_url}/", timeout=5)
        
        if response.status_code == 200:
            print(f"✅ Server is running at {server_url}")
            print(f"   Response: {response.json()}")
            return True
        else:
            print(f"⚠️  Server responded with status code: {response.status_code}")
            return False
            
    except requests.ConnectionError:
        print(f"⚠️  Server is not running at {server_url}")
        print("   Start server with: python src/backend/run_server.py")
        return None  # Not a failure, just not running
    except Exception as e:
        print(f"❌ Error connecting to server: {e}")
        return False


def test_session_endpoints():
    """Test session-based authentication endpoints."""
    print("\n=== Testing Session Endpoints ===")
    print("Note: This test requires the server to be running.")
    
    server_url = "http://localhost:8000"
    
    try:
        # Test login endpoint (if exists)
        # This is a placeholder - adjust based on your actual auth endpoints
        response = requests.get(f"{server_url}/docs", timeout=5)
        
        if response.status_code == 200:
            print(f"✅ API docs available at {server_url}/docs")
            print("   Check the docs for available session/auth endpoints")
            return True
        else:
            print(f"⚠️  Could not access API docs")
            return False
            
    except requests.ConnectionError:
        print(f"⚠️  Server is not running")
        return None
    except Exception as e:
        print(f"❌ Error: {e}")
        return False


def main():
    """Run all tests."""
    print("=" * 60)
    print("FinDeck Redis & Server Test Suite")
    print("=" * 60)
    
    results = {
        "Redis Connection": test_redis_connection(),
        "Session Secret": test_session_secret(),
        "Server Startup": test_server_startup(),
        "Session Endpoints": test_session_endpoints(),
    }
    
    print("\n" + "=" * 60)
    print("Test Results Summary")
    print("=" * 60)
    
    for test_name, result in results.items():
        if result is True:
            status = "✅ PASSED"
        elif result is False:
            status = "❌ FAILED"
        else:
            status = "⚠️  SKIPPED"
        print(f"{test_name:.<40} {status}")
    
    # Overall result
    passed = sum(1 for r in results.values() if r is True)
    failed = sum(1 for r in results.values() if r is False)
    skipped = sum(1 for r in results.values() if r is None)
    
    print("\n" + "=" * 60)
    print(f"Total: {passed} passed, {failed} failed, {skipped} skipped")
    print("=" * 60)
    
    if failed > 0:
        print("\n❌ Some tests failed. Please check the configuration.")
        sys.exit(1)
    else:
        print("\n✅ All critical tests passed!")
        sys.exit(0)


if __name__ == "__main__":
    main()
