"""
Quick test to generate a professional AI Pro PPT like the ones in examples/professional_demo
"""
from pathlib import Path
from src.converter.excel_to_ppt_converter import ExcelToPPTConverter

def generate_professional_ppt():
    """Generate a professional AI Pro tier presentation"""
    
    # Use one of your existing Excel files
    excel_files = [
        "examples/Portfolio Allocation Data.xlsx",
        "examples/Risk Metrics Data.xlsx",
        "examples/Sample_pnl.xlsx"
    ]
    
    # Check which file exists
    excel_file = None
    for file_path in excel_files:
        if Path(file_path).exists():
            excel_file = file_path
            break
    
    if not excel_file:
        print("❌ No Excel file found. Please provide an Excel file.")
        return
    
    print(f"📊 Using Excel file: {excel_file}")
    print("🤖 Generating ENHANCED AI PRO tier presentation with 8+ slides...")
    print("   ✨ Executive Summary - High-level insights")
    print("   ✨ Key Metrics - 6 major KPIs")
    print("   ✨ Data Insights - Detailed analysis")
    print("   ✨ Sector Distribution - Category breakdown")
    print("   ✨ Key Data Insights - Deep dive analysis")
    print("   ✨ Top Performers - Ranked list")
    print("   ✨ Trend Analysis - Growth patterns")
    print("   ✨ Professional design & branding")
    
    # Create converter with AI PRO tier
    converter = ExcelToPPTConverter(
        user_tier='ai_pro',  # Use AI PRO for professional look
        user_id='test_user'
    )
    
    # Output file
    output_file = f"test_professional_ai_pro.pptx"
    
    try:
        # Convert with ENHANCED professional structure (8+ slides)
        result = converter.convert_professional(
            excel_path=excel_file,
            output_path=output_file,
            presentation_title=f"Professional Analysis - {Path(excel_file).stem}",
            template_name="corporate_blue",  # Use professional template
            use_professional_structure=True  # Force enhanced 8+ slide structure
        )
        
        print(f"\n✅ SUCCESS! ENHANCED Professional PPT generated: {output_file}")
        print(f"📊 Slides created: {result.get('slides_created', 'N/A')}")
        print(f"📄 Includes: Title, Executive Summary, Key Metrics, Data Insights,")
        print(f"            Sector Distribution, Key Data Insights, Top Performers,")
        print(f"            Trend Analysis, and Closing slides")
        print(f"🎨 Template used: {result.get('template_used', 'corporate_blue')}")
        print(f"🤖 AI features used: {', '.join(result.get('ai_features_used', []))}")
        print(f"\n📂 Open the file to see the comprehensive professional presentation!")
        
        # Open the file
        import os
        os.startfile(output_file)
        
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    generate_professional_ppt()
