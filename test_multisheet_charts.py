"""
Multi-Sheet Chart Generator for Company Bundle
Generates one comprehensive PowerPoint with charts from ALL sheets
"""

import sys
import os
sys.path.insert(0, 'src')

from converter.excel_reader import excel_reader, get_sheet_names, excel_reader_all_sheets
from converter.ppt_writer import create_presentation, create_auto_chart_slide
from converter.chart_detector import should_create_chart, detect_chart_type

def create_multi_sheet_presentation(excel_path: str, output_path: str, 
                                   max_charts: int = 20, 
                                   sheet_filter: str = None):
    """
    Create a PowerPoint presentation with charts from multiple sheets.
    
    Args:
        excel_path: Path to Excel file with multiple sheets
        output_path: Where to save the PowerPoint
        max_charts: Maximum number of charts to create
        sheet_filter: Only process sheets containing this keyword (e.g., "Prices", "Income")
    """
    
    print("="*70)
    print(f"📊 MULTI-SHEET CHART GENERATOR")
    print("="*70)
    print(f"📁 Input: {excel_path}")
    print(f"💾 Output: {output_path}")
    
    # Get all sheet names
    all_sheets = get_sheet_names(excel_path)
    print(f"\n📋 Found {len(all_sheets)} sheets in Excel file")
    
    # Filter sheets if requested
    if sheet_filter:
        sheets_to_process = [s for s in all_sheets if sheet_filter.lower() in s.lower()]
        print(f"🔍 Filtering for '{sheet_filter}': {len(sheets_to_process)} sheets")
    else:
        sheets_to_process = all_sheets
    
    # Create presentation
    prs = create_presentation(
        title="Multi-Company Financial Analysis",
        subtitle=f"Data from {len(sheets_to_process)} sheets"
    )
    
    charts_created = 0
    sheets_processed = 0
    
    print(f"\n{'='*70}")
    print("🔄 Processing Sheets...")
    print(f"{'='*70}\n")
    
    for i, sheet_name in enumerate(sheets_to_process, 1):
        if charts_created >= max_charts:
            print(f"⚠️  Reached max charts limit ({max_charts}). Stopping.")
            break
        
        try:
            print(f"{i}. 📄 Sheet: '{sheet_name}'")
            
            # Read sheet data
            df = excel_reader(excel_path, sheet=sheet_name)
            
            if df is None or df.empty:
                print(f"   ⚠️  Empty sheet, skipping")
                continue
            
            print(f"   📊 Data: {df.shape[0]} rows × {df.shape[1]} columns")
            
            # For large datasets (e.g., price data), use recent data only
            if df.shape[0] > 100:
                df = df.tail(50)  # Last 50 rows
                print(f"   🔄 Using last 50 rows for visualization")
            
            # Check if suitable for charting
            if not should_create_chart(df):
                print(f"   ⚠️  Not suitable for charting (too many/few rows or no numeric data)")
                continue
            
            # Detect chart type
            chart_type, config = detect_chart_type(df)
            
            if chart_type is None:
                print(f"   ⚠️  No appropriate chart type detected")
                continue
            
            print(f"   ✅ Chart Type: {chart_type.upper()}")
            print(f"   📝 Config: {config}")
            
            # Create chart slide
            slide = create_auto_chart_slide(prs, df, title=f"{sheet_name}")
            
            if slide:
                charts_created += 1
                print(f"   ✅ Chart #{charts_created} created!")
            else:
                print(f"   ❌ Failed to create chart")
            
            sheets_processed += 1
            
        except Exception as e:
            print(f"   ❌ Error: {str(e)}")
            continue
    
    # Save presentation
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    
    print(f"\n{'='*70}")
    print(f"✅ COMPLETE!")
    print(f"{'='*70}")
    print(f"📊 Charts Created: {charts_created}")
    print(f"📄 Sheets Processed: {sheets_processed}/{len(sheets_to_process)}")
    print(f"💾 Saved to: {output_path}")
    print(f"{'='*70}\n")
    
    return output_path


def create_company_specific_charts(excel_path: str, company: str, output_path: str):
    """
    Create charts for a specific company (e.g., AAPL, MSFT, GOOGL, AMZN, TSLA)
    """
    
    print("="*70)
    print(f"📊 COMPANY-SPECIFIC CHARTS: {company}")
    print("="*70)
    
    # Get all sheets for this company
    all_sheets = get_sheet_names(excel_path)
    company_sheets = [s for s in all_sheets if s.startswith(company)]
    
    print(f"📋 Found {len(company_sheets)} sheets for {company}:")
    for sheet in company_sheets:
        print(f"   • {sheet}")
    
    # Create presentation
    prs = create_presentation(
        title=f"{company} Financial Analysis",
        subtitle="Comprehensive Data Visualization"
    )
    
    charts_created = 0
    
    print(f"\n{'='*70}")
    print("🔄 Processing Company Data...")
    print(f"{'='*70}\n")
    
    for sheet_name in company_sheets:
        try:
            print(f"📄 {sheet_name}")
            df = excel_reader(excel_path, sheet=sheet_name)
            
            if df is None or df.empty:
                print(f"   ⚠️  Empty, skipping\n")
                continue
            
            # For price data, limit to recent data for better visualization
            if 'Prices' in sheet_name and df.shape[0] > 50:
                df = df.tail(30)  # Last 30 trading days
                print(f"   📊 Using last 30 days ({df.shape[0]} rows)")
            else:
                print(f"   📊 {df.shape[0]} rows × {df.shape[1]} columns")
            
            if should_create_chart(df):
                chart_type, config = detect_chart_type(df)
                if chart_type:
                    print(f"   ✅ Creating {chart_type.upper()} chart")
                    slide = create_auto_chart_slide(prs, df, title=sheet_name)
                    if slide:
                        charts_created += 1
                        print(f"   ✅ Chart #{charts_created} created!\n")
                    else:
                        print(f"   ❌ Failed\n")
                else:
                    print(f"   ⚠️  No chart type detected\n")
            else:
                print(f"   ⚠️  Not suitable for charting\n")
                
        except Exception as e:
            print(f"   ❌ Error: {str(e)}\n")
    
    # Save
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    
    print(f"{'='*70}")
    print(f"✅ {company} Analysis Complete!")
    print(f"📊 Charts: {charts_created}")
    print(f"💾 Saved: {output_path}")
    print(f"{'='*70}\n")
    
    return output_path


def demo_multi_sheet_capabilities():
    """Demonstrate multi-sheet chart generation capabilities"""
    
    excel_file = "examples/Company_Data/company_bundle.xlsx"
    
    print("\n" + "="*70)
    print("🚀 TESTING MULTI-SHEET CHART CAPABILITIES")
    print("="*70 + "\n")
    
    # Test 1: All sheets (limited to 15 charts)
    print("TEST 1: Generate presentation with charts from ALL sheets")
    print("-" * 70)
    output1 = create_multi_sheet_presentation(
        excel_file,
        "examples/demo_PPT/All_Companies_Multi_Sheet.pptx",
        max_charts=15
    )
    
    # Test 2: Only Price sheets
    print("\nTEST 2: Generate presentation with only PRICE charts")
    print("-" * 70)
    output2 = create_multi_sheet_presentation(
        excel_file,
        "examples/demo_PPT/All_Companies_Prices_Only.pptx",
        max_charts=10,
        sheet_filter="Prices"
    )
    
    # Test 3: Only Income sheets
    print("\nTEST 3: Generate presentation with only INCOME charts")
    print("-" * 70)
    output3 = create_multi_sheet_presentation(
        excel_file,
        "examples/demo_PPT/All_Companies_Income_Only.pptx",
        max_charts=10,
        sheet_filter="Income"
    )
    
    # Test 4: Company-specific (AAPL)
    print("\nTEST 4: Generate presentation for AAPL only")
    print("-" * 70)
    output4 = create_company_specific_charts(
        excel_file,
        "AAPL",
        "examples/demo_PPT/AAPL_Complete_Analysis.pptx"
    )
    
    # Test 5: Company-specific (TSLA)
    print("\nTEST 5: Generate presentation for TSLA only")
    print("-" * 70)
    output5 = create_company_specific_charts(
        excel_file,
        "TSLA",
        "examples/demo_PPT/TSLA_Complete_Analysis.pptx"
    )
    
    print("\n" + "="*70)
    print("🎉 ALL TESTS COMPLETE!")
    print("="*70)
    print("\n📁 Generated Presentations:")
    print("   1. All_Companies_Multi_Sheet.pptx - Mixed charts from all sheets")
    print("   2. All_Companies_Prices_Only.pptx - Only price trend charts")
    print("   3. All_Companies_Income_Only.pptx - Only income statement charts")
    print("   4. AAPL_Complete_Analysis.pptx - All AAPL data charts")
    print("   5. TSLA_Complete_Analysis.pptx - All TSLA data charts")
    print("\n💡 Check examples/demo_PPT/ folder to view all presentations!")
    print("="*70 + "\n")


if __name__ == "__main__":
    demo_multi_sheet_capabilities()
