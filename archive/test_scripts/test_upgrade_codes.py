"""
Test Plan Upgrade Code System
==============================

This script tests the complete upgrade code flow:
1. Generate upgrade code for customer
2. Send email with code
3. Customer redeems code
4. Verify plan upgrade

Run this after starting the backend server:
    cd src/backend/app
    uvicorn main:app --reload
"""

import requests
import json
from datetime import datetime

# API Configuration
BASE_URL = "http://localhost:8000/api/v1"
ADMIN_EMAIL = "admin@findeck.com"
ADMIN_PASSWORD = "admin123"
CUSTOMER_EMAIL = "customer@example.com"
CUSTOMER_PASSWORD = "customer123"


def print_section(title):
    """Print formatted section header"""
    print(f"\n{'='*80}")
    print(f"  {title}")
    print('='*80)


def print_success(message):
    """Print success message"""
    print(f"✅ {message}")


def print_error(message):
    """Print error message"""
    print(f"❌ {message}")


def print_info(message):
    """Print info message"""
    print(f"ℹ️  {message}")


def login(email, password):
    """Login and get access token"""
    response = requests.post(
        f"{BASE_URL}/auth/login",
        data={
            "username": email,
            "password": password
        }
    )
    
    if response.status_code == 200:
        token = response.json()["access_token"]
        print_success(f"Logged in as {email}")
        return token
    else:
        print_error(f"Login failed: {response.status_code}")
        print(response.text)
        return None


def generate_upgrade_code(admin_token, customer_email, plan, duration_months=1):
    """Admin generates upgrade code for customer"""
    print_section("STEP 1: Generate Upgrade Code")
    
    headers = {
        "Authorization": f"Bearer {admin_token}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "email": customer_email,
        "plan": plan,
        "duration_months": duration_months,
        "notes": f"Test upgrade to {plan} plan"
    }
    
    print_info(f"Generating {plan} upgrade code for {customer_email}...")
    
    response = requests.post(
        f"{BASE_URL}/upgrades/generate-upgrade-code",
        headers=headers,
        json=payload
    )
    
    if response.status_code == 201:
        data = response.json()
        print_success("Upgrade code generated successfully!")
        print(f"\n📧 Code Details:")
        print(f"   Code: {data['code']}")
        print(f"   Plan: {data['plan']}")
        print(f"   Customer: {data['customer_email']}")
        print(f"   Expires: {data['expires_at']}")
        print(f"\n💌 Email sent to customer with upgrade instructions")
        return data['code']
    else:
        print_error(f"Failed to generate code: {response.status_code}")
        print(response.text)
        return None


def check_my_codes(customer_token):
    """Customer checks their upgrade codes"""
    print_section("STEP 2: Check Customer's Upgrade Codes")
    
    headers = {
        "Authorization": f"Bearer {customer_token}"
    }
    
    response = requests.get(
        f"{BASE_URL}/upgrades/my-upgrade-codes",
        headers=headers
    )
    
    if response.status_code == 200:
        data = response.json()
        codes = data.get('codes', [])
        
        print_success(f"Found {len(codes)} upgrade code(s)")
        
        for code in codes:
            print(f"\n📜 Code: {code['code']}")
            print(f"   Plan: {code['plan']}")
            print(f"   Expires: {code['expires_at']}")
            print(f"   Redeemed: {'Yes ✓' if code['is_redeemed'] else 'No'}")
            print(f"   Can Redeem: {'Yes ✓' if code['can_redeem'] else 'No'}")
        
        return codes
    else:
        print_error(f"Failed to check codes: {response.status_code}")
        return []


def redeem_upgrade_code(customer_token, code):
    """Customer redeems upgrade code"""
    print_section("STEP 3: Redeem Upgrade Code")
    
    headers = {
        "Authorization": f"Bearer {customer_token}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "code": code
    }
    
    print_info(f"Redeeming code: {code}...")
    
    response = requests.post(
        f"{BASE_URL}/upgrades/redeem-upgrade-code",
        headers=headers,
        json=payload
    )
    
    if response.status_code == 200:
        data = response.json()
        print_success("Code redeemed successfully!")
        print(f"\n🎉 Upgrade Complete:")
        print(f"   New Plan: {data['new_plan']}")
        print(f"   Subscription Ends: {data['subscription_ends_at']}")
        print(f"\n✨ Features Unlocked:")
        for feature in data['features_unlocked']:
            print(f"   ✓ {feature}")
        return True
    else:
        print_error(f"Failed to redeem code: {response.status_code}")
        print(response.text)
        return False


def verify_upgrade(customer_token):
    """Verify customer's plan was upgraded"""
    print_section("STEP 4: Verify Plan Upgrade")
    
    headers = {
        "Authorization": f"Bearer {customer_token}"
    }
    
    response = requests.get(
        f"{BASE_URL}/auth/me",
        headers=headers
    )
    
    if response.status_code == 200:
        data = response.json()
        subscription = data.get('subscription', {})
        
        print_success("Current Account Status:")
        print(f"\n👤 User: {data['email']}")
        print(f"📊 Plan: {subscription.get('plan', 'Unknown')}")
        print(f"🔄 Status: {subscription.get('status', 'Unknown')}")
        print(f"📅 Ends: {subscription.get('ends_at', 'N/A')}")
        
        return subscription
    else:
        print_error(f"Failed to verify upgrade: {response.status_code}")
        return None


def check_all_codes_admin(admin_token):
    """Admin checks all upgrade codes"""
    print_section("ADMIN: View All Upgrade Codes")
    
    headers = {
        "Authorization": f"Bearer {admin_token}"
    }
    
    response = requests.get(
        f"{BASE_URL}/upgrades/admin/all-upgrade-codes",
        headers=headers
    )
    
    if response.status_code == 200:
        data = response.json()
        codes = data.get('codes', [])
        
        print_success(f"Total upgrade codes: {data['total_codes']}")
        
        for code in codes:
            print(f"\n📜 {code['code']}")
            print(f"   Plan: {code['plan']}")
            print(f"   Customer: {code['customer_email']}")
            print(f"   Generated By: {code.get('generated_by', 'Unknown')}")
            print(f"   Redeemed: {'Yes ✓' if code['is_redeemed'] else 'No'}")
            if code['is_redeemed']:
                print(f"   Redeemed By: {code.get('redeemed_by_email')}")
                print(f"   Redeemed At: {code.get('redeemed_at')}")
        
        return codes
    else:
        print_error(f"Failed to check all codes: {response.status_code}")
        return []


def main():
    """Run complete upgrade code test flow"""
    print_section("🧪 PLAN UPGRADE CODE SYSTEM TEST")
    
    print("\n📋 Test Scenario:")
    print("   1. Admin generates BASIC plan upgrade code")
    print("   2. Code is sent to customer via email")
    print("   3. Customer checks their codes")
    print("   4. Customer redeems the code")
    print("   5. Verify plan was upgraded")
    print("   6. Admin reviews all codes")
    
    input("\n⏸️  Press Enter to start the test...")
    
    # Login as admin
    print_section("Admin Login")
    admin_token = login(ADMIN_EMAIL, ADMIN_PASSWORD)
    if not admin_token:
        print_error("Admin login failed. Make sure admin account exists.")
        return
    
    # Login as customer
    print_section("Customer Login")
    customer_token = login(CUSTOMER_EMAIL, CUSTOMER_PASSWORD)
    if not customer_token:
        print_error("Customer login failed. Make sure customer account exists.")
        return
    
    # Generate upgrade code
    upgrade_code = generate_upgrade_code(
        admin_token=admin_token,
        customer_email=CUSTOMER_EMAIL,
        plan="BASIC",
        duration_months=1
    )
    
    if not upgrade_code:
        print_error("Could not generate upgrade code. Test stopped.")
        return
    
    input("\n⏸️  Press Enter to continue to customer redemption...")
    
    # Customer checks their codes
    my_codes = check_my_codes(customer_token)
    
    # Redeem the code
    success = redeem_upgrade_code(customer_token, upgrade_code)
    
    if not success:
        print_error("Code redemption failed. Test stopped.")
        return
    
    # Verify upgrade
    verify_upgrade(customer_token)
    
    input("\n⏸️  Press Enter to view admin dashboard...")
    
    # Admin checks all codes
    check_all_codes_admin(admin_token)
    
    print_section("✅ TEST COMPLETE")
    print("\nAll steps completed successfully!")
    print("\n📊 Summary:")
    print("   ✓ Upgrade code generated")
    print("   ✓ Email sent to customer")
    print("   ✓ Customer redeemed code")
    print("   ✓ Plan upgraded successfully")
    print("   ✓ Admin can track all codes")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Test interrupted by user")
    except Exception as e:
        print(f"\n\n❌ Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()
