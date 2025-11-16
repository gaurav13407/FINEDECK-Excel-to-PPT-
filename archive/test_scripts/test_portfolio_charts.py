"""
Test Enhanced Charts with Portfolio Allocation Data
This should show clear PIE charts due to allocation/percentage data
"""

import sys
import os
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / 'src'))

from pptx import Presentation
from converter.enhanced_professional_builder import EnhancedProfessionalBuilder
from converter.excel_reader import excel_reader_all_sheets

def test_portfolio_conversion():
    """Test with Portfolio Allocation Data - should show PIE charts!"""
    
    print("\n" + "="*80)
    print("🎯 TESTING WITH PORTFOLIO ALLOCATION DATA (Should Show PIE CHARTS!)")
    print("="*80 + "\n")
    
    # Use Portfolio Allocation Data
    excel_file = "examples/Portfolio Allocation Data.xlsx"
    
    if not os.path.exists(excel_file):
        print(f"❌ File not found: {excel_file}")
        return False
    
    print(f"📂 Reading: {excel_file}")
    
    # Read Excel data
    sheets_dict = excel_reader_all_sheets(excel_file)
    sheets_data = [(name, df) for name, df in sheets_dict.items() if df is not None]
    
    print(f"✅ Read {len(sheets_data)} sheets")
    
    # Show data structure
    if sheets_data:
        print("\n📊 First sheet data:")
        df = sheets_data[0][1]
        print(f"  Rows: {len(df)}")
        print(f"  Columns: {df.columns.tolist()}")
        print(f"  Data types: {df.dtypes.to_dict()}")
        print(f"\n  First 3 rows:")
        print(df.head(3))
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = 9144000
    prs.slide_height = 6858000
    
    # Royal Purple template
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
    
    print(f"\n🎨 Using template: {template['name']}")
    
    # Create builder
    builder = EnhancedProfessionalBuilder(
        ai_service=None,
        user_metadata={'name': 'Test User'},
        user_tier='ai_pro',
        use_finance_charts=True
    )
    
    print(f"\n🏗️  Building presentation with ENHANCED CHARTS...")
    print(f"📊 Expected: PIE charts for allocation data\n")
    
    # Build presentation
    results = builder.build_presentation(
        prs,
        sheets_data,
        project_name="Portfolio Allocation Analysis",
        template=template
    )
    
    # Save output
    output_path = "test_output/portfolio_with_pie_charts.pptx"
    os.makedirs("test_output", exist_ok=True)
    prs.save(output_path)
    
    print("\n" + "="*80)
    print("📊 RESULTS")
    print("="*80)
    print(f"✅ Slides created: {results['slides_created']}")
    print(f"✅ Output saved: {output_path}")
    
    print("\n" + "="*80)
    print("🔍 WHAT TO LOOK FOR:")
    print("="*80)
    print("1. Open: test_output/portfolio_with_pie_charts.pptx")
    print("2. Check Slide 6 (Sector Distribution) - Should have PIE CHART in PURPLE")
    print("3. Charts should use Royal Purple color scheme")
    print("4. Look for percentage labels on pie slices")
    print("="*80 + "\n")
    
    return True

if __name__ == "__main__":
    test_portfolio_conversion()
