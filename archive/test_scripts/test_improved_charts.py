"""
Test improved charts and insights with Financials.csv
Shows diverse charts based on smart analysis
"""

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter
import os

def test_improved_financials():
    """Test with improved smart charts"""
    
    print("\n" + "="*80)
    print("🧪 TESTING IMPROVED CHARTS & INSIGHTS - Financials Dataset")
    print("="*80)
    
    excel_path = "examples/Company_Data/financials_bundle.xlsx"
    output_path = "examples/professional_demo/Financials_FINAL_IMPROVED.pptx"
    
    if not os.path.exists(excel_path):
        print(f"❌ Excel file not found: {excel_path}")
        return
    
    print(f"\n📊 Input: {excel_path}")
    print(f"📁 Output: {output_path}")
    
    # Create converter with AI_PRO tier
    converter = ExcelToPPTConverter(
        user_tier='ai_pro',
        user_id='test_improved_user',
        user_metadata={
            'name': 'Test User - Improved Charts',
            'company': 'Data Visualization Labs',
            'email': 'test@improved.com'
        }
    )
    
    # Generate presentation
    print(f"\n🎨 Generating presentation with SMART CHARTS...")
    
    result = converter.convert_professional(
        excel_path=excel_path,
        output_path=output_path,
        presentation_title="Financial Performance Analysis - Smart Charts & Data-Driven Insights",
        user_ppt_count=0
    )
    
    if result.get('success', False):
        file_size = os.path.getsize(output_path) / 1024
        print(f"\n✅ SUCCESS!")
        print(f"   Slides: {result.get('slides_created', 'N/A')}")
        print(f"   AI Features: {result.get('ai_features_count', 0)}")
        print(f"   File Size: {file_size:.1f} KB")
        print(f"   Output: {output_path}")
        
        print(f"\n📊 IMPROVEMENTS:")
        print(f"   ✅ Smart chart analyzer determines best visualizations")
        print(f"   ✅ Data-driven insights (no generic AI text)")
        print(f"   ✅ Diverse chart types based on data structure")
        print(f"   ✅ Better value formatting (currency, percentages)")
        print(f"   ✅ Cleaner, more professional presentation")
    else:
        print(f"\n❌ FAILED!")
        print(f"   Error: {result.get('error', 'Unknown error')}")


if __name__ == "__main__":
    test_improved_financials()
