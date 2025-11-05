"""
Simple test - generate ONE professional PPT for BASIC tier
"""

import sys
import os
from pathlib import Path

# Add project root
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter

print("=" * 80)
print("🎨 SINGLE TIER TEST - BASIC")
print("=" * 80)

# Setup
excel_file = "examples/Sample_pnl.xlsx"
output_dir = "examples/professional_demo"
os.makedirs(output_dir, exist_ok=True)

# User metadata
user_metadata = {
    'name': 'John Doe',
    'company': 'FinDeck Analytics Inc.',
    'email': 'john.doe@findeck.com'
}

print(f"\n📁 Input: {excel_file}")
print(f"📂 Output: {output_dir}")
print(f"👤 User: {user_metadata['name']}")
print(f"🎯 Tier: BASIC")

try:
    # Create converter
    converter = ExcelToPPTConverter(
        user_tier='basic',
        user_id='test_123',
        user_metadata=user_metadata
    )
    
    output_file = os.path.join(output_dir, "test_basic_professional.pptx")
    
    print(f"\n🔄 Converting...")
    
    # Convert
    result = converter.convert_professional(
        excel_path=excel_file,
        output_path=output_file,
        presentation_title="Q4 Financial Performance Report",
        user_ppt_count=0
    )
    
    if result['success']:
        file_size = os.path.getsize(output_file) / 1024
        
        print(f"\n✅ SUCCESS!")
        print(f"   📊 Slides: {result['slides_created']}")
        print(f"   🎨 Template: {result['template_used']}")
        print(f"   🤖 AI Features: {result.get('ai_features_used', [])}")
        print(f"   📦 Size: {file_size:.1f} KB")
        print(f"   💾 File: {output_file}")
        
        if result.get('errors'):
            print(f"\n   ⚠️  Errors encountered:")
            for error in result['errors']:
                print(f"      - {error}")
    else:
        print(f"\n❌ FAILED: {result.get('error')}")

except Exception as e:
    print(f"\n❌ ERROR: {str(e)}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 80)
