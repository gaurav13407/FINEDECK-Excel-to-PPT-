"""
Test Tiered Conversion with AI Features for AI_PRO
===================================================
This script tests the full conversion logic with all AI features:
- AI Titles
- AI Summaries
- AI Insights (5 bullets)
- AI Layout Optimization
- AI Chart Recommendations
- Smart Template Selection
- Multiple Charts with Legends
"""

import requests
import json
from pathlib import Path

# Configuration
BASE_URL = "http://localhost:8000/api/v1"
EXCEL_FILE = "examples/Financials.xlsx"  # Using your Financials data
TEMPLATE = "dark_finance"  # Dark finance theme for AI_PRO

def test_ai_pro_conversion():
    print("="*80)
    print("🚀 TESTING AI_PRO TIER CONVERSION")
    print("="*80)
    
    # Step 1: Login
    print("\n📝 Step 1: Logging in...")
    email = input("Enter email (default: gaurav13407@outlook.com): ").strip() or "gaurav13407@outlook.com"
    password = input("Enter password: ")
    
    login_url = f"{BASE_URL}/auth/login"
    login_response = requests.post(login_url, json={"email": email, "password": password})
    
    if login_response.status_code != 200:
        print(f"❌ Login failed: {login_response.text}")
        return
    
    token = login_response.json().get("access_token")
    print(f"✅ Login successful!")
    
    headers = {
        "Authorization": f"Bearer {token}"
    }
    
    # Step 2: Check tier features
    print("\n🎯 Step 2: Checking AI_PRO features...")
    features_url = f"{BASE_URL}/tiered-convert/tier-features"
    features_response = requests.get(features_url, headers=headers)
    
    if features_response.status_code == 200:
        features = features_response.json()
        print(f"✅ Current Plan: {features['tier'].upper()}")
        print(f"   Features Available:")
        for feature, enabled in features['features'].items():
            status = "✅" if enabled else "❌"
            print(f"     {status} {feature}")
        print(f"   Templates: {', '.join(features['templates'])}")
        print(f"   PPTs Used: {features['usage']['presentations_this_month']}/{features['usage']['presentations_limit']}")
    
    # Step 3: Upload and convert with all AI features
    print(f"\n🔄 Step 3: Converting Excel to PPT with AI_PRO features...")
    print(f"   Excel File: {EXCEL_FILE}")
    print(f"   Template: {TEMPLATE}")
    
    # Check if file exists
    excel_path = Path(EXCEL_FILE)
    if not excel_path.exists():
        print(f"❌ Excel file not found: {EXCEL_FILE}")
        print(f"   Please update the EXCEL_FILE variable in the script")
        return
    
    convert_url = f"{BASE_URL}/tiered-convert/tiered-convert"
    
    with open(excel_path, 'rb') as f:
        files = {'file': (excel_path.name, f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
        data = {
            'template_name': TEMPLATE,
            'presentation_title': 'AI-Powered Financial Analysis'
        }
        
        print(f"\n⏳ Converting... (this may take 30-60 seconds with AI processing)")
        response = requests.post(convert_url, headers=headers, files=files, data=data)
    
    if response.status_code == 200:
        # Save the PowerPoint file
        output_filename = "AI_PRO_Demo_Output.pptx"
        with open(output_filename, 'wb') as f:
            f.write(response.content)
        
        # Get metadata from headers
        slides_created = response.headers.get('X-Slides-Created', 'Unknown')
        template_used = response.headers.get('X-Template-Used', 'Unknown')
        ai_features = response.headers.get('X-AI-Features', '').split(',')
        ai_tokens = response.headers.get('X-AI-Tokens', '0')
        ai_cost = response.headers.get('X-AI-Cost', '0.00')
        
        print(f"\n🎉 SUCCESS! Presentation created!")
        print(f"="*80)
        print(f"\n📊 Conversion Summary:")
        print(f"   Output File: {output_filename}")
        print(f"   Slides Created: {slides_created}")
        print(f"   Template Used: {template_used}")
        print(f"   AI Features Applied: {', '.join(filter(None, ai_features))}")
        print(f"   AI Tokens Used: {ai_tokens}")
        print(f"   AI Cost: ${ai_cost}")
        
        print(f"\n✨ AI_PRO Features Included:")
        print(f"   📝 AI-Generated Titles - Smart, context-aware slide titles")
        print(f"   📄 AI Summaries - Executive summaries for each sheet")
        print(f"   💡 AI Insights - 5 key insights from your data")
        print(f"   📊 Smart Charts - Auto-detected chart types with legends")
        print(f"   🎨 Dark Finance Theme - Professional financial styling")
        print(f"   📈 Chart Analysis - AI-recommended visualizations")
        
        print(f"\n🎯 What to Look For:")
        print(f"   1. Title slide with AI-generated title")
        print(f"   2. Executive Summary with AI insights")
        print(f"   3. Multiple chart types (bar, line, pie, table)")
        print(f"   4. Color-coded legends on all charts")
        print(f"   5. Dark finance theme colors")
        print(f"   6. AI-generated summaries for each section")
        
        print(f"\n📂 Open the file to see all AI features in action!")
        print(f"   File location: {Path(output_filename).absolute()}")
        
    else:
        print(f"\n❌ Conversion failed: {response.status_code}")
        try:
            error_detail = response.json()
            print(f"   Error: {error_detail.get('detail', response.text)}")
        except:
            print(f"   Error: {response.text}")
    
    print(f"\n" + "="*80)

if __name__ == "__main__":
    print("\n🔥 AI_PRO TIER CONVERSION TEST")
    print("This will test all advanced AI features including:")
    print("  • AI Titles, Summaries, & Insights")
    print("  • Smart Chart Detection")
    print("  • Dark Finance Theme")
    print("  • Multiple Chart Types with Legends")
    print()
    
    test_ai_pro_conversion()
