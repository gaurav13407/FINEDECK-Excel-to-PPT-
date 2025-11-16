"""
Test Tiered Conversion API Endpoints
Tests all API endpoints for the tiered conversion system
"""

import requests
import json
from pathlib import Path

# Configuration
BASE_URL = "http://localhost:8000/api/v1"
TEST_EXCEL = "examples/Sample_pnl.xlsx"

# Test user tokens (you'll need to get these from actual login)
# For now, we'll test with placeholder tokens
TEST_TOKENS = {
    "free": "YOUR_FREE_USER_TOKEN",
    "basic": "YOUR_BASIC_USER_TOKEN",
    "pro": "YOUR_PRO_USER_TOKEN",
    "ai_pro": "YOUR_AI_PRO_USER_TOKEN"
}


def test_tier_features(token: str, tier_name: str):
    """Test GET /tiered/tier-features endpoint"""
    print(f"\n{'='*80}")
    print(f"Testing Tier Features - {tier_name.upper()}")
    print(f"{'='*80}")
    
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(f"{BASE_URL}/tiered/tier-features", headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        print(f"✅ Tier: {data['tier_name']} (${data['price']}/month)")
        print(f"   PPT Limit: {data['ppt_limit']}")
        print(f"   Max Sheets: {data['max_sheets']}")
        print(f"   Templates: {len(data['templates_available']) if isinstance(data['templates_available'], list) else 'All'}")
        print(f"\n   AI Features:")
        for feature, enabled in data['ai_features'].items():
            status = "✅" if enabled else "❌"
            print(f"   {status} {feature}")
        
        if data.get('upgrade_options'):
            print(f"\n   Available Upgrades:")
            for upgrade in data['upgrade_options']:
                print(f"   → {upgrade['name']} (${upgrade['price']}/month)")
    else:
        print(f"❌ Failed: {response.status_code} - {response.text}")


def test_usage_stats(token: str, tier_name: str):
    """Test GET /tiered/usage-stats endpoint"""
    print(f"\n{'='*80}")
    print(f"Testing Usage Stats - {tier_name.upper()}")
    print(f"{'='*80}")
    
    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(f"{BASE_URL}/tiered/usage-stats", headers=headers)
    
    if response.status_code == 200:
        data = response.json()
        print(f"✅ Current Month: {data['current_month']}")
        print(f"   Tier: {data['tier']}")
        print(f"   PPTs Created: {data['ppt_created']}/{data['ppt_limit']}")
        print(f"   Remaining: {data['ppt_remaining']}")
        print(f"   Usage: {data['usage_percentage']:.1f}%")
        print(f"   AI Tokens Used: {data['ai_tokens_used']}")
        print(f"   AI Cost: ${data['ai_cost_total']:.6f}")
    else:
        print(f"❌ Failed: {response.status_code} - {response.text}")


def test_preview_ai(token: str, tier_name: str):
    """Test POST /tiered/preview-ai endpoint"""
    print(f"\n{'='*80}")
    print(f"Testing AI Preview - {tier_name.upper()}")
    print(f"{'='*80}")
    
    if not Path(TEST_EXCEL).exists():
        print(f"❌ Test file not found: {TEST_EXCEL}")
        return
    
    headers = {"Authorization": f"Bearer {token}"}
    
    with open(TEST_EXCEL, 'rb') as f:
        files = {'file': (Path(TEST_EXCEL).name, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        response = requests.post(f"{BASE_URL}/tiered/preview-ai", headers=headers, files=files)
    
    if response.status_code == 200:
        data = response.json()
        print(f"✅ File: {data['filename']}")
        print(f"   Sheets: {data['sheets_found']}")
        print(f"   Tier: {data['tier']}")
        print(f"   AI Features: {', '.join(data['ai_features_available'])}")
        
        for i, rec in enumerate(data['recommendations'], 1):
            print(f"\n   Sheet {i}: {rec['sheet_name']} ({rec['rows']} rows, {rec['columns']} cols)")
            
            if 'ai_title' in rec:
                print(f"   📝 AI Title: {rec['ai_title']}")
            
            if 'ai_template' in rec:
                template = rec['ai_template']
                if 'auto_selected' in template:
                    print(f"   🎨 AI Template: {template['auto_selected']}")
            
            if 'ai_chart' in rec:
                chart = rec['ai_chart']
                if 'recommended_chart' in chart:
                    print(f"   📊 AI Chart: {chart['recommended_chart']} (confidence: {chart.get('confidence', 0):.2f})")
        
        if 'ai_usage' in data:
            usage = data['ai_usage']
            print(f"\n   AI Usage:")
            print(f"   Tokens: {usage.get('total_tokens', 0)}")
            print(f"   Cost: ${usage.get('total_cost', 0):.6f}")
    
    elif response.status_code == 402:
        print(f"⚠️  Upgrade Required: {response.json().get('detail')}")
    else:
        print(f"❌ Failed: {response.status_code} - {response.text}")


def test_conversion(token: str, tier_name: str):
    """Test POST /tiered/tiered-convert endpoint"""
    print(f"\n{'='*80}")
    print(f"Testing Conversion - {tier_name.upper()}")
    print(f"{'='*80}")
    
    if not Path(TEST_EXCEL).exists():
        print(f"❌ Test file not found: {TEST_EXCEL}")
        return
    
    headers = {"Authorization": f"Bearer {token}"}
    
    with open(TEST_EXCEL, 'rb') as f:
        files = {'file': (Path(TEST_EXCEL).name, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        data = {'presentation_title': f'Test {tier_name} Conversion'}
        response = requests.post(f"{BASE_URL}/tiered/tiered-convert", headers=headers, files=files, data=data)
    
    if response.status_code == 200:
        # Save the PPT file
        output_file = f"test_api_{tier_name}_output.pptx"
        with open(output_file, 'wb') as f:
            f.write(response.content)
        
        # Print headers
        print(f"✅ Conversion successful!")
        print(f"   Output: {output_file}")
        print(f"   Slides: {response.headers.get('X-Slides-Created', 'unknown')}")
        print(f"   Template: {response.headers.get('X-Template-Used', 'unknown')}")
        print(f"   AI Features: {response.headers.get('X-AI-Features', 'none')}")
        print(f"   AI Tokens: {response.headers.get('X-AI-Tokens', '0')}")
        print(f"   AI Cost: ${response.headers.get('X-AI-Cost', '0')}")
    
    elif response.status_code == 402:
        print(f"⚠️  Limit Reached: {response.json().get('detail')}")
    else:
        print(f"❌ Failed: {response.status_code}")
        try:
            print(f"   Error: {response.json().get('detail', response.text)}")
        except:
            print(f"   Error: {response.text}")


def main():
    """Run all tests"""
    print("="*80)
    print("TIERED CONVERSION API TESTS")
    print("="*80)
    print("\n⚠️  NOTE: You need to update TEST_TOKENS with real authentication tokens")
    print("   Get tokens by logging in users with different subscription tiers\n")
    
    # Test each tier
    for tier_name, token in TEST_TOKENS.items():
        if token == f"YOUR_{tier_name.upper()}_USER_TOKEN":
            print(f"\n⏭️  Skipping {tier_name} (no token configured)")
            continue
        
        try:
            # Test tier features
            test_tier_features(token, tier_name)
            
            # Test usage stats
            test_usage_stats(token, tier_name)
            
            # Test AI preview (if tier supports it)
            if tier_name != "free":
                test_preview_ai(token, tier_name)
            
            # Test conversion
            test_conversion(token, tier_name)
            
        except Exception as e:
            print(f"❌ Error testing {tier_name}: {str(e)}")
    
    print("\n" + "="*80)
    print("🎉 ALL TESTS COMPLETE!")
    print("="*80)


if __name__ == "__main__":
    # Quick test without authentication (will fail but shows endpoints)
    print("\n" + "="*80)
    print("QUICK ENDPOINT CHECK (without authentication)")
    print("="*80)
    
    endpoints = [
        ("GET", "/tiered/tier-features"),
        ("GET", "/tiered/usage-stats"),
        ("POST", "/tiered/preview-ai"),
        ("POST", "/tiered/tiered-convert")
    ]
    
    print("\n📋 Available Endpoints:")
    for method, endpoint in endpoints:
        print(f"   {method:6s} {BASE_URL}{endpoint}")
    
    print("\n💡 To test with authentication:")
    print("   1. Start the backend server: uvicorn main:app --reload")
    print("   2. Create test users with different subscription tiers")
    print("   3. Login and get authentication tokens")
    print("   4. Update TEST_TOKENS in this script")
    print("   5. Run: python test_api_endpoints.py")
    
    print("\n📝 Example: Create test users via API")
    print("   POST /api/v1/auth/register")
    print("   Body: {\"name\": \"Test User\", \"email\": \"test@example.com\", \"password\": \"password123\"}")
    print("\n   Then update subscription tier in database:")
    print("   db.users.updateOne({email: 'test@example.com'}, {$set: {'subscription.plan': 'ai_pro'}})")
