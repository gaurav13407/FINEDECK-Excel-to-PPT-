"""
Real-World Demo: Comprehensive Financial Analysis for AAPL
Creates a professional multi-chart presentation from AAPL_Financial_Data.xlsx
"""

import sys
import os
sys.path.insert(0, 'src')

from converter.excel_reader import excel_reader, get_sheet_names
from converter.ppt_writer import create_presentation, create_auto_chart_slide
from converter.chart_detector import should_create_chart, detect_chart_type

def create_aapl_financial_presentation():
    """
    Create a comprehensive financial presentation for AAPL.
    This is a REAL-WORLD EXAMPLE showing how the system handles actual financial data.
    """
    
    excel_file = "example/Company_Data/AAPL_Financial_Data.xlsx"
    output_file = "examples/demo_PPT/AAPL_Financial_Analysis_Complete.pptx"
    
    print("="*70)
    print("📊 REAL-WORLD DEMO: APPLE (AAPL) FINANCIAL ANALYSIS")
    print("="*70)
    print(f"📁 Input: {excel_file}")
    print(f"💾 Output: {output_file}\n")
    
    # Get all sheets
    sheets = get_sheet_names(excel_file)
    print(f"📋 Found {len(sheets)} sheets in AAPL financial data:")
    for i, sheet in enumerate(sheets, 1):
        print(f"   {i}. {sheet}")
    
    # Create presentation
    prs = create_presentation(
        title="Apple Inc. (AAPL) Financial Analysis",
        subtitle="Comprehensive Multi-Sheet Data Visualization"
    )
    
    print(f"\n{'='*70}")
    print("🔄 Processing Each Sheet...")
    print(f"{'='*70}\n")
    
    charts_created = 0
    
    for sheet_name in sheets:
        print(f"📄 Sheet: '{sheet_name}'")
        print("-" * 70)
        
        try:
            # Read sheet
            df = excel_reader(excel_file, sheet=sheet_name)
            
            if df is None or df.empty:
                print(f"   ⚠️  Empty sheet\n")
                continue
            
            print(f"   📊 Data: {df.shape[0]} rows × {df.shape[1]} columns")
            print(f"   📋 Columns: {df.columns.tolist()[:5]}{'...' if len(df.columns) > 5 else ''}")
            
            # Show sample data
            if df.shape[0] > 0:
                print(f"\n   📝 Sample Data (first 3 rows):")
                sample = df.head(3)
                for idx, row in sample.iterrows():
                    row_dict = {k: v for k, v in row.items() if k and str(k).strip()}
                    if row_dict:
                        # Show first 3 columns only
                        items = list(row_dict.items())[:3]
                        print(f"      Row {idx+1}: {dict(items)}")
            
            # For large datasets (price history), use recent data
            original_rows = df.shape[0]
            if df.shape[0] > 100:
                df = df.tail(50)
                print(f"\n   🔄 Sampling: Using last 50 rows (from {original_rows} total)")
            
            # Check if suitable for charting
            if not should_create_chart(df):
                print(f"   ⚠️  Not suitable for charting")
                print(f"      Reason: May have too many/few rows or insufficient numeric data\n")
                continue
            
            # Detect chart type
            chart_type, config = detect_chart_type(df)
            
            if chart_type is None:
                print(f"   ⚠️  No appropriate chart type detected")
                print(f"      This sheet might need manual chart configuration\n")
                continue
            
            print(f"\n   ✅ Chart Type Detected: {chart_type.upper()}")
            print(f"   📝 Configuration:")
            for key, value in config.items():
                if isinstance(value, list):
                    print(f"      • {key}: {', '.join(str(v) for v in value[:3])}")
                else:
                    print(f"      • {key}: {value}")
            
            # Create chart
            slide = create_auto_chart_slide(prs, df, title=f"AAPL - {sheet_name}")
            
            if slide:
                charts_created += 1
                print(f"\n   🎉 SUCCESS! Chart #{charts_created} created for '{sheet_name}'")
            else:
                print(f"\n   ❌ Failed to create chart")
            
        except Exception as e:
            print(f"   ❌ Error processing sheet: {str(e)}")
        
        print()  # Blank line between sheets
    
    # Save presentation
    os.makedirs(os.path.dirname(output_file), exist_ok=True)
    prs.save(output_file)
    
    # Final summary
    print(f"{'='*70}")
    print(f"✅ PRESENTATION COMPLETE!")
    print(f"{'='*70}")
    print(f"\n📊 Statistics:")
    print(f"   • Total Sheets Processed: {len(sheets)}")
    print(f"   • Charts Created: {charts_created}")
    print(f"   • Success Rate: {(charts_created/len(sheets)*100):.1f}%")
    print(f"\n💾 Saved to: {output_file}")
    print(f"\n🎯 Presentation Structure:")
    print(f"   • Slide 1: Title Slide")
    for i in range(charts_created):
        print(f"   • Slide {i+2}: Data Visualization Chart")
    print(f"   • Total Slides: {charts_created + 1}")
    
    print(f"\n{'='*70}")
    print(f"📂 Open the file to see your professional financial analysis!")
    print(f"{'='*70}\n")
    
    return output_file


def analyze_sheet_details():
    """
    Detailed analysis of each sheet to show what the system sees
    """
    excel_file = "example/Company_Data/AAPL_Financial_Data.xlsx"
    sheets = get_sheet_names(excel_file)
    
    print("\n" + "="*70)
    print("🔍 DETAILED SHEET ANALYSIS")
    print("="*70 + "\n")
    
    for sheet_name in sheets:
        print(f"📄 {sheet_name}")
        print("-" * 70)
        
        try:
            df = excel_reader(excel_file, sheet=sheet_name)
            
            if df is None or df.empty:
                print("   Status: Empty\n")
                continue
            
            # Analyze data structure
            numeric_cols = df.select_dtypes(include=['int64', 'float64']).columns.tolist()
            categorical_cols = df.select_dtypes(include=['object', 'string']).columns.tolist()
            
            print(f"   Rows: {df.shape[0]}")
            print(f"   Columns: {df.shape[1]}")
            print(f"   Numeric columns: {len(numeric_cols)}")
            print(f"   Categorical columns: {len(categorical_cols)}")
            
            # Sample data
            if df.shape[0] > 0:
                print(f"\n   First row preview:")
                first_row = df.iloc[0]
                for col in df.columns[:5]:  # Show first 5 columns
                    val = first_row[col]
                    if pd.notna(val):
                        print(f"      {col}: {val}")
            
            # Chart potential
            working_df = df.tail(50) if df.shape[0] > 100 else df
            if should_create_chart(working_df):
                chart_type, config = detect_chart_type(working_df)
                if chart_type:
                    print(f"\n   ✅ Chart Potential: {chart_type.upper()}")
                else:
                    print(f"\n   ⚠️  Chart Potential: Unclear")
            else:
                print(f"\n   ❌ Chart Potential: Not suitable")
            
        except Exception as e:
            print(f"   Error: {str(e)}")
        
        print()


if __name__ == "__main__":
    import pandas as pd
    
    # Main demo
    output = create_aapl_financial_presentation()
    
    # Detailed analysis
    print("\n" + "="*70)
    print("Would you like to see detailed sheet analysis? (y/n)")
    print("="*70)
    # Uncomment to run analysis:
    # analyze_sheet_details()
    
    print(f"\n🎉 Demo Complete! Open {output} to view your presentation!")
