"""
Test Login and Redeem Flow
===========================
This script tests the full authentication and redeem flow.
"""

import requests
import json
from getpass import getpass

BASE_URL = "http://localhost:8000/api/v1"

def test_login_and_redeem():
    print("="*80)
    print("🧪 TESTING LOGIN AND REDEEM FLOW")
    print("="*80)
    
    # Step 1: Login
    print("\n📝 Step 1: Login")
    email = input("Enter email (default: gaurav13407@outlook.com): ").strip() or "gaurav13407@outlook.com"
    password = getpass("Enter password: ")
    
    login_url = f"{BASE_URL}/auth/login"
    login_data = {
        "email": email,
        "password": password
    }
    
    print(f"🔐 Attempting login for: {email}")
    try:
        response = requests.post(login_url, json=login_data)
        
        if response.status_code == 200:
            data = response.json()
            token = data.get("access_token")
            print(f"✅ Login successful!")
            print(f"   Token (first 50 chars): {token[:50]}...")
        else:
            print(f"❌ Login failed: {response.status_code}")
            print(f"   Response: {response.text}")
            return
    except Exception as e:
        print(f"❌ Error during login: {e}")
        return
    
    # Step 2: Test /auth/me endpoint
    print("\n👤 Step 2: Testing /auth/me endpoint")
    me_url = f"{BASE_URL}/auth/me"
    headers = {
        "Authorization": f"Bearer {token}"
    }
    
    try:
        response = requests.get(me_url, headers=headers)
        if response.status_code == 200:
            user_data = response.json()
            print(f"✅ User info retrieved!")
            print(f"   Email: {user_data.get('email')}")
            print(f"   Plan: {user_data.get('subscription', {}).get('plan', 'N/A')}")
        else:
            print(f"❌ Failed to get user info: {response.status_code}")
            print(f"   Response: {response.text}")
            return
    except Exception as e:
        print(f"❌ Error getting user info: {e}")
        return
    
    # Step 3: Redeem upgrade code
    print("\n🎁 Step 3: Testing upgrade code redemption")
    code = input("Enter upgrade code (default: AI_PRO-SWVIPE94XJKF): ").strip() or "AI_PRO-SWVIPE94XJKF"
    
    redeem_url = f"{BASE_URL}/upgrades/redeem-upgrade-code"
    redeem_data = {
        "code": code
    }
    
    print(f"🔑 Attempting to redeem: {code}")
    try:
        response = requests.post(redeem_url, headers=headers, json=redeem_data)
        
        if response.status_code == 200:
            result = response.json()
            print(f"✅ Redemption successful!")
            print(f"   Message: {result.get('message')}")
            print(f"   New Plan: {result.get('new_plan')}")
            print(f"   Previous Plan: {result.get('previous_plan')}")
        else:
            print(f"❌ Redemption failed: {response.status_code}")
            print(f"   Response: {response.text}")
            return
    except Exception as e:
        print(f"❌ Error during redemption: {e}")
        return
    
    # Step 4: Verify plan updated
    print("\n✅ Step 4: Verifying plan update")
    try:
        response = requests.get(me_url, headers=headers)
        if response.status_code == 200:
            user_data = response.json()
            new_plan = user_data.get('subscription', {}).get('plan', 'N/A')
            print(f"✅ Plan verified!")
            print(f"   Current Plan: {new_plan}")
            
            if new_plan.lower() == "ai_pro":
                print(f"\n🎉 SUCCESS! Your plan has been upgraded to AI_PRO!")
            else:
                print(f"\n⚠️  Plan is {new_plan}, expected ai_pro")
        else:
            print(f"❌ Failed to verify: {response.status_code}")
    except Exception as e:
        print(f"❌ Error during verification: {e}")
    
    print("\n" + "="*80)
    print("✅ TEST COMPLETE")
    print("="*80)

if __name__ == "__main__":
    test_login_and_redeem()
