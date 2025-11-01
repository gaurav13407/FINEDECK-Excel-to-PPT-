"""
Quick visual summary of what charts were generated
"""

import sys
import os
sys.path.insert(0, 'src')

from converter.excel_reader import excel_reader
from converter.chart_detector import detect_chart_type

print("="*70)
print("📊 CHART GENERATION SUMMARY")
print("="*70)

samples = [
    ("Portfolio Allocation Data.xlsx", None, "🥧 PIE CHART"),
    ("Risk Metrics Data.xlsx", None, "📊 BAR CHART"),
    ("Sample_pnl.xlsx", "Sheet1", "📈 LINE CHART")
]

for filename, sheet, expected in samples:
    filepath = f"examples/{filename}"
    df = excel_reader(filepath, sheet=sheet)
    chart_type, config = detect_chart_type(df)
    
    print(f"\n{'='*70}")
    print(f"📁 File: {filename}")
    if sheet:
        print(f"   Sheet: {sheet}")
    print(f"{'='*70}")
    print(f"📊 Data Shape: {df.shape[0]} rows × {df.shape[1]} columns")
    print(f"📋 Columns: {', '.join(df.columns.tolist())}")
    print(f"🎯 Detected: {chart_type.upper() if chart_type else 'NONE'}")
    print(f"✨ Expected: {expected}")
    print(f"✅ Match: {'YES' if chart_type in expected.lower() else 'NO'}")
    
    if config:
        print(f"\n📝 Chart Config:")
        for key, value in config.items():
            print(f"   • {key}: {value}")
    
    print(f"\n📊 Sample Data:")
    print(df.head(3).to_string(index=False))

print(f"\n{'='*70}")
print(f"✅ ALL PRESENTATIONS GENERATED IN: examples/demo_PPT/")
print(f"{'='*70}")
print("\n🎉 Files to check:")
print("   1. CHART_DEMO.pptx - All chart types in one presentation")
print("   2. Portfolio_Report_with_Chart.pptx - Pie chart example")
print("   3. Risk_Metrics_with_Chart.pptx - Bar chart example")
print("   4. PnL_Trend_with_Chart.pptx - Line chart example")
print("\n💡 Open these files in PowerPoint to see the results!")
