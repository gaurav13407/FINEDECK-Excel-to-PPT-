"""
Demo script to test intelligent chart generation with all sample files
"""

import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from converter.excel_reader import excel_reader
from converter.ppt_writer import df_to_ppt, create_presentation, create_auto_chart_slide
from converter.chart_detector import detect_chart_type

def demo_chart_generation():
    """Generate a comprehensive demo PowerPoint with all chart types"""
    
    print("="*70)
    print("📊 INTELLIGENT CHART GENERATION DEMO")
    print("="*70)
    
    # Create presentation
    prs = create_presentation(
        title="FinDeck Chart Demo", 
        subtitle="Intelligent Chart Detection & Generation"
    )
    
    # Sample 1: Portfolio Allocation (should create PIE chart)
    print("\n1️⃣  Portfolio Allocation Data → Expecting: PIE CHART")
    df1 = excel_reader('examples/Portfolio Allocation Data.xlsx')
    print(f"   Data: {df1.shape[0]} rows, {df1.shape[1]} columns")
    print(f"   Columns: {df1.columns.tolist()}")
    chart_type, config = detect_chart_type(df1)
    print(f"   ✅ Detected: {chart_type.upper()} chart")
    print(f"   Config: {config}")
    create_auto_chart_slide(prs, df1, title="Portfolio Allocation by Asset")
    
    # Sample 2: Risk Metrics (should create BAR chart)
    print("\n2️⃣  Risk Metrics Data → Expecting: BAR CHART")
    df2 = excel_reader('examples/Risk Metrics Data.xlsx')
    print(f"   Data: {df2.shape[0]} rows, {df2.shape[1]} columns")
    print(f"   Columns: {df2.columns.tolist()}")
    chart_type, config = detect_chart_type(df2)
    print(f"   ✅ Detected: {chart_type.upper()} chart")
    print(f"   Config: {config}")
    create_auto_chart_slide(prs, df2, title="Risk Metrics Comparison")
    
    # Sample 3: PnL Data (should create LINE chart)
    print("\n3️⃣  Sample PnL Data → Expecting: LINE CHART")
    df3 = excel_reader('examples/Sample_pnl.xlsx', sheet='Sheet1')
    print(f"   Data: {df3.shape[0]} rows, {df3.shape[1]} columns")
    print(f"   Columns: {df3.columns.tolist()}")
    chart_type, config = detect_chart_type(df3)
    print(f"   ✅ Detected: {chart_type.upper()} chart")
    print(f"   Config: {config}")
    create_auto_chart_slide(prs, df3, title="Revenue Trend Over Quarters")
    
    # Sample 4: Multi-metric comparison (should create COLUMN chart)
    print("\n4️⃣  Multi-Metric Comparison → Expecting: COLUMN CHART")
    # Use the same PnL data but show all metrics together
    create_auto_chart_slide(prs, df3, title="Revenue, Cost & Profit Comparison")
    
    # Save demo presentation
    output_path = "examples/demo_PPT/CHART_DEMO.pptx"
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    prs.save(output_path)
    
    print("\n" + "="*70)
    print(f"✅ DEMO COMPLETE!")
    print(f"📁 Saved to: {output_path}")
    print("="*70)
    
    return output_path


def demo_full_reports():
    """Generate full reports with auto mode"""
    
    print("\n" + "="*70)
    print("📊 GENERATING FULL REPORTS WITH CHARTS")
    print("="*70)
    
    # Report 1: Portfolio with chart + details
    print("\n📈 Portfolio Report (auto mode with charts)...")
    df1 = excel_reader('examples/Portfolio Allocation Data.xlsx')
    out1 = df_to_ppt(
        df1, 
        out_path="examples/demo_PPT/Portfolio_Report_with_Chart.pptx",
        title="Portfolio Allocation Report",
        subtitle="Automated with Chart Generation",
        mode="auto",
        include_charts=True
    )
    print(f"   ✅ Saved: {out1}")
    
    # Report 2: Risk Metrics
    print("\n📊 Risk Metrics Report...")
    df2 = excel_reader('examples/Risk Metrics Data.xlsx')
    out2 = df_to_ppt(
        df2,
        out_path="examples/demo_PPT/Risk_Metrics_with_Chart.pptx",
        title="Risk Metrics Analysis",
        subtitle="With Automatic Chart Detection",
        mode="auto",
        include_charts=True
    )
    print(f"   ✅ Saved: {out2}")
    
    # Report 3: PnL Trend
    print("\n📈 PnL Trend Report...")
    df3 = excel_reader('examples/Sample_pnl.xlsx', sheet='Sheet1')
    out3 = df_to_ppt(
        df3,
        out_path="examples/demo_PPT/PnL_Trend_with_Chart.pptx",
        title="Profit & Loss Trend",
        subtitle="Quarterly Analysis with Charts",
        mode="auto",
        include_charts=True
    )
    print(f"   ✅ Saved: {out3}")
    
    print("\n" + "="*70)
    print("✅ ALL REPORTS GENERATED!")
    print("="*70)


if __name__ == "__main__":
    # Generate demo with all chart types
    demo_chart_generation()
    
    # Generate full reports
    demo_full_reports()
    
    print("\n🎉 Check the 'examples/demo_PPT' folder for all generated presentations!")
