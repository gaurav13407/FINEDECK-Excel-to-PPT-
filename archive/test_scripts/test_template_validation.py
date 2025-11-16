"""
Simulate backend template processing to debug the issue
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter

def test_template_selection():
    print("\n" + "="*80)
    print("🔍 TESTING TEMPLATE SELECTION LOGIC")
    print("="*80)
    
    # Create converter like backend does
    converter = ExcelToPPTConverter(
        user_tier='ai_pro',
        user_id='test_user',
        user_metadata={'name': 'Test User', 'email': 'test@test.com', 'company': 'Test Co'}
    )
    
    print(f"\n✅ Converter created for tier: ai_pro")
    
    # Get allowed templates
    allowed = converter.get_allowed_templates()
    print(f"\n📋 Allowed templates ({len(allowed)}):")
    for i, tmpl in enumerate(allowed, 1):
        marker = " ← FIRST (default)" if i == 1 else ""
        marker2 = " ← THIS IS ROYAL_PURPLE!" if tmpl == 'royal_purple' else ""
        print(f"   {i}. {tmpl}{marker}{marker2}")
    
    # Test validation
    template_name = 'royal_purple'
    print(f"\n🎨 Testing validation for: {template_name}")
    print(f"   Is in allowed list? {template_name in allowed}")
    
    if template_name not in allowed:
        print(f"   ❌ Template NOT in allowed list - would use default: {allowed[0]}")
    else:
        print(f"   ✅ Template IS in allowed list - would use: {template_name}")
    
    # Test what the converter method would do
    print(f"\n🧪 SIMULATING CONVERT_PROFESSIONAL LOGIC:")
    print(f"   Input template_name: '{template_name}'")
    
    # Simulate the logic from convert_professional
    if template_name is None or template_name == 'None' or template_name == '':
        print(f"   ⚠️ Template is None/empty, using default: {allowed[0]}")
        final_template = allowed[0]
    elif template_name not in allowed:
        print(f"   ⚠️ Template not in allowed list, using default: {allowed[0]}")
        final_template = allowed[0]
    else:
        print(f"   ✅ Template is valid, using: {template_name}")
        final_template = template_name
    
    print(f"\n🎯 FINAL RESULT: {final_template}")
    
    if final_template != template_name:
        print(f"\n❌ PROBLEM DETECTED!")
        print(f"   Expected: {template_name}")
        print(f"   Got: {final_template}")
        print(f"\n💡 This means the validation logic is failing!")
    else:
        print(f"\n✅ Template selection working correctly!")

if __name__ == "__main__":
    test_template_selection()
