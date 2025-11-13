"""
Test Backend Integration - Data Intelligence Engine
Tests the full flow: Analysis → Excel Reader → PPT Writer
"""

import sys
sys.path.insert(0, r'c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src')
sys.path.insert(0, r'c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\src\backend\app')

from converter.excel_reader import excel_reader
from converter.ppt_writer import df_to_ppt
from services.data_intelligence import DataIntelligenceEngine
import tempfile
import os

# File to test
excel_file = r'c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\DV+Sales+Data.xlsx'

print("🔍 Step 1: Analyzing Excel file with Data Intelligence Engine...")
engine = DataIntelligenceEngine()
analysis = engine.analyze_file(excel_file)

print(f"✅ Analysis Complete!")
print(f"   Records: {analysis['executive_summary']['total_records']:,}")
print(f"   Columns: {analysis['executive_summary']['total_columns']}")
print(f"   Numeric Metrics: {analysis['executive_summary']['numeric_columns']}")

print("\n📖 Step 2: Reading Excel data with excel_reader...")
df = excel_reader(excel_file, sheet=0)
print(f"✅ Data loaded: {len(df)} rows x {len(df.columns)} columns")

print("\n🎨 Step 3: Creating PowerPoint with df_to_ppt...")
output_ppt = 'DV_Sales_Backend_Test.pptx'

# Create intelligent title
intelligent_title = f"Intelligent Data Analysis - {analysis['executive_summary']['total_records']:,} Records"
intelligent_subtitle = f"Auto-Generated | {analysis['executive_summary']['numeric_columns']} Metrics | {analysis['executive_summary']['categorical_columns']} Categories"

df_to_ppt(
    df=df,
    out_path=output_ppt,
    title=intelligent_title,
    subtitle=intelligent_subtitle,
    title_col=None,
    mode="table",
    limit=50  # Limit to 50 rows for reasonable PPT size
)

print(f"✅ PowerPoint Created: {output_ppt}")
print("\n" + "="*80)
print("🎯 BACKEND INTEGRATION TEST SUCCESSFUL!")
print("="*80)
print("The intelligent conversion endpoint will:")
print("  1. ✅ Analyze data with Intelligence Engine")
print("  2. ✅ Read Excel with excel_reader")  
print("  3. ✅ Generate PPT with df_to_ppt")
print("  4. ✅ Return downloadable file")

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
