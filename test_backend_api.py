"""
Test the actual backend API endpoint to see what it returns
"""

import requests
import os

# Backend URL
API_URL = "http://localhost:8000/api/v1/tiered/tiered-convert"

# Test file
excel_file = r"C:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\examples\Portfolio Allocation Data.xlsx"

# You need to get your auth token from localStorage in browser
# For now, let's just check if backend is responding
print("🔍 Testing Backend API Response Headers...")
print("="*80)

# Prepare the request
files = {
    'file': ('Portfolio Allocation Data.xlsx', open(excel_file, 'rb'), 
             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
}

data = {
    'template_name': 'royal_purple',
    'use_finance_charts': 'false'
}

# You'll need to add your auth token here
headers = {
    # 'Authorization': 'Bearer YOUR_TOKEN_HERE'
}

print("\n📤 Request:")
print(f"   URL: {API_URL}")
print(f"   File: {os.path.basename(excel_file)}")
print(f"   Template: royal_purple")
print(f"   Use Finance Charts: false")

print("\n⚠️  NOTE: This will fail without auth token")
print("   To get your auth token:")
print("   1. Open browser DevTools (F12)")
print("   2. Go to Console tab")
print("   3. Type: localStorage.getItem('authToken')")
print("   4. Copy the token")
print("   5. Add it to this script in the headers")

print("\n" + "="*80)
print("💡 Instead, let's check the backend terminal logs when you upload!")
print("="*80)
print("\nWhen you upload through the browser, check the backend terminal for:")
print("   ✓ 'slides_created': 10")
print("   ✓ 'template_used': 'royal_purple'")
print("   ✓ 'ai_features_used': [...]")
print("\nIf you see these in the backend logs but not in browser,")
print("then it's a header transmission issue.")

# Close file
files['file'][1].close()
