"""
Test Backend Integration of Enhancement Features
Tests that enhanced features work through the actual backend converter
"""

import sys
import os
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from pptx import Presentation
from converter.enhanced_professional_builder import EnhancedProfessionalBuilder
from converter.excel_reader import excel_reader_all_sheets
import pandas as pd

def test_backend_integration():
    """Test that enhancements work through EnhancedProfessionalBuilder"""
    
    print("\n" + "="*80)
    print("🧪 TESTING BACKEND INTEGRATION OF ENHANCEMENT FEATURES")
    print("="*80 + "\n")
    
    # Use sample Excel file
    excel_file = "examples/Sample_pnl.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ Test file not found: {excel_file}")
        return False
    
    print(f"📂 Reading Excel file: {excel_file}")
    
    # Read Excel data
    sheets_dict = excel_reader_all_sheets(excel_file)
    sheets_data = [(name, df) for name, df in sheets_dict.items() if df is not None]
    
    print(f"✅ Read {len(sheets_data)} sheets")
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = 9144000  # 10 inches
    prs.slide_height = 6858000  # 7.5 inches
    
    # Load template (Royal Purple for visibility)
    template = {
        'name': 'Royal Purple',
        'colors': {
            'primary': '#4A148C',
            'secondary': '#7B1FA2',
            'accent': '#CE93D8',
            'text': '#212121',
            'light': '#F3E5F5',
            'white': '#FFFFFF',
            'chart_colors': ['#4A148C', '#7B1FA2', '#9C27B0', '#BA68C8', '#CE93D8', '#E1BEE7']
        }
    }
    
    print(f"🎨 Using template: {template['name']}")
    
    # Create builder with AI_PRO tier for all features
    builder = EnhancedProfessionalBuilder(
        ai_service=None,
        user_metadata={'name': 'Test User'},
        user_tier='ai_pro',
        use_finance_charts=True
    )
    
    print(f"🏗️  Building presentation...")
    
    # Build presentation
    results = builder.build_presentation(
        prs,
        sheets_data,
        project_name="Backend Integration Test",
        template=template
    )
    
    # Save output
    output_path = "test_output/backend_integration_test.pptx"
    os.makedirs("test_output", exist_ok=True)
    prs.save(output_path)
    
    print("\n" + "="*80)
    print("📊 RESULTS")
    print("="*80)
    print(f"✅ Slides created: {results['slides_created']}")
    print(f"✅ Slide names: {', '.join(results['slide_names'])}")
    print(f"✅ AI features used: {', '.join(results['ai_features_used'])}")
    print(f"✅ Output saved: {output_path}")
    
    # Verify enhancements are present
    print("\n" + "="*80)
    print("🔍 VERIFICATION")
    print("="*80)
    
    expected_features = [
        'enhanced_cover_slide',
        'ai_insights_slide'
    ]
    
    all_present = True
    for feature in expected_features:
        if feature in results['ai_features_used']:
            print(f"✅ {feature}: INTEGRATED")
        else:
            print(f"❌ {feature}: MISSING")
            all_present = False
    
    # Check for specific slide names
    expected_slides = ['Enhanced Cover Slide', 'AI Insights']
    for slide_name in expected_slides:
        if slide_name in results['slide_names']:
            print(f"✅ {slide_name}: PRESENT")
        else:
            print(f"❌ {slide_name}: MISSING")
            all_present = False
    
    print("\n" + "="*80)
    if all_present:
        print("✅ ✅ ✅ ALL ENHANCEMENT FEATURES SUCCESSFULLY INTEGRATED! ✅ ✅ ✅")
    else:
        print("❌ SOME FEATURES ARE MISSING - CHECK INTEGRATION")
    print("="*80 + "\n")
    
    return all_present

if __name__ == "__main__":
    success = test_backend_integration()
    sys.exit(0 if success else 1)
