"""
Template Testing Script
=======================
This script tests ALL 10 templates by generating a PowerPoint for each one.
It will help verify if the template system is working correctly.

Usage:
    python test_all_templates.py

Output:
    Creates 10 PPT files in the 'test_output' folder, one for each template.
"""

import sys
import os
from pathlib import Path
import json

# Add src to path
script_dir = Path(__file__).parent
src_dir = script_dir / 'src'
sys.path.insert(0, str(src_dir))

from converter.enhanced_professional_builder import EnhancedProfessionalBuilder
from pptx import Presentation

# All 10 templates to test
TEMPLATES_TO_TEST = [
    'corporate_blue',
    'dark_finance',
    'elegant_gray',
    'forest_green',
    'minimal_white',
    'modern_tech',
    'ocean_blue',
    'royal_purple',
    'sunset_orange',
    'vibrant_gradient'
]

def load_template_json(template_name):
    """Load template JSON file"""
    template_path = script_dir / 'src' / 'templates' / 'built_in' / f'{template_name}.json'
    
    if not template_path.exists():
        print(f"❌ Template file not found: {template_path}")
        return None
    
    with open(template_path, 'r') as f:
        template_data = json.load(f)
    
    print(f"✅ Loaded template: {template_data.get('name', template_name)}")
    return template_data

def create_sample_data():
    """Create sample Excel data for testing - must be list of tuples (sheet_name, dataframe)"""
    import pandas as pd
    
    # Create DataFrames for each sheet
    df1 = pd.DataFrame({
        'Metric': ['Revenue', 'Expenses', 'Profit'],
        'Q1': [100000, 70000, 30000],
        'Q2': [120000, 80000, 40000],
        'Q3': [135000, 85000, 50000],
        'Q4': [150000, 90000, 60000]
    })
    
    df2 = pd.DataFrame({
        'Metric': ['Customer Growth', 'Market Share', 'Satisfaction'],
        'Value': ['25%', '15%', '92%'],
        'Change': ['+5%', '+2%', '+3%']
    })
    
    # Return as list of tuples (sheet_name, dataframe)
    return [
        ('Financial Summary', df1),
        ('Key Metrics', df2)
    ]

def test_template(template_name, output_dir):
    """Test a single template"""
    print(f"\n{'='*60}")
    print(f"🎨 Testing Template: {template_name.upper()}")
    print(f"{'='*60}")
    
    # Load template
    template_data = load_template_json(template_name)
    if not template_data:
        return False
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = 9144000
    prs.slide_height = 6858000
    
    # Create builder - use basic tier to skip AI insights that cause errors
    builder = EnhancedProfessionalBuilder(
        ai_service=None,
        user_metadata={'tier': 'basic'},
        user_tier='basic',
        use_finance_charts=False
    )
    
    # Get sample data
    sheets_data = create_sample_data()
    
    # Build presentation with template
    print(f"📋 Building presentation with template...")
    builder.build_presentation(
        prs=prs,
        sheets_data=sheets_data,
        project_name=f"Template Test - {template_data.get('name', template_name)}",
        template=template_data
    )
    
    # Save presentation
    output_file = output_dir / f"{template_name}_test.pptx"
    prs.save(str(output_file))
    
    print(f"✅ Saved: {output_file}")
    print(f"📊 Slides created: {len(prs.slides)}")
    
    # Verify template colors were applied
    if hasattr(builder, 'template_colors'):
        primary_color = builder.template_colors.get('navy', 'Unknown')
        accent_color = builder.template_colors.get('light_blue', 'Unknown')
        print(f"🎨 Primary color (navy): {primary_color}")
        print(f"🎨 Accent color (light_blue): {accent_color}")
    
    return True

def main():
    """Main test function"""
    print("\n" + "="*60)
    print("🧪 TEMPLATE SYSTEM TEST")
    print("="*60)
    print(f"Testing {len(TEMPLATES_TO_TEST)} templates...")
    
    # Create output directory
    output_dir = script_dir / 'test_output'
    output_dir.mkdir(exist_ok=True)
    print(f"📁 Output directory: {output_dir}")
    
    # Test each template
    results = {}
    for template_name in TEMPLATES_TO_TEST:
        try:
            success = test_template(template_name, output_dir)
            results[template_name] = 'PASS' if success else 'FAIL'
        except Exception as e:
            print(f"❌ ERROR testing {template_name}: {e}")
            import traceback
            traceback.print_exc()
            results[template_name] = 'ERROR'
    
    # Print summary
    print("\n" + "="*60)
    print("📊 TEST SUMMARY")
    print("="*60)
    
    passed = sum(1 for r in results.values() if r == 'PASS')
    failed = sum(1 for r in results.values() if r == 'FAIL')
    errors = sum(1 for r in results.values() if r == 'ERROR')
    
    for template_name, result in results.items():
        icon = '✅' if result == 'PASS' else '❌'
        print(f"{icon} {template_name:20s} - {result}")
    
    print(f"\n📈 Results: {passed} passed, {failed} failed, {errors} errors")
    print(f"📁 Output files in: {output_dir}")
    
    if passed == len(TEMPLATES_TO_TEST):
        print("\n🎉 ALL TEMPLATES PASSED!")
        print("\n📝 Next steps:")
        print("   1. Open each .pptx file in test_output/")
        print("   2. Verify the colors match the template name")
        print("   3. Check title slide, content slides, and charts")
        return 0
    else:
        print(f"\n⚠️  {failed + errors} templates failed")
        return 1

if __name__ == '__main__':
    sys.exit(main())
