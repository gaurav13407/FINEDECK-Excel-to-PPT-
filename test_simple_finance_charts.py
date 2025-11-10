"""
Test the SimpleFinanceChartBuilder integration
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from pptx import Presentation
from pptx.util import Inches
import pandas as pd
from src.converter.simple_finance_charts import SimpleFinanceChartBuilder

# Create test data
print("📊 Creating test data...")

# Test 1: Portfolio allocation (should create PIE chart)
allocation_data = pd.DataFrame({
    'Sector': ['Technology', 'Healthcare', 'Finance', 'Energy', 'Consumer'],
    'Allocation': [35, 25, 20, 12, 8]
})

# Test 2: Quarterly revenue (should create LINE chart)
quarterly_data = pd.DataFrame({
    'Quarter': ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024'],
    'Revenue': [120, 145, 156, 178],
    'Profit': [25, 32, 38, 45]
})

# Test 3: Top performers (should create COLUMN chart)
performance_data = pd.DataFrame({
    'Company': ['AAPL', 'MSFT', 'GOOGL', 'AMZN', 'TSLA'],
    'Performance': [15.2, 12.8, 10.5, 9.3, 8.1]
})

# Create presentation
print("📄 Creating presentation...")
prs = Presentation()

# Create chart builder
chart_builder = SimpleFinanceChartBuilder()

# Test 1: PIE chart
print("\n🥧 Test 1: Creating PIE chart from allocation data...")
slide1 = prs.slides.add_slide(prs.slide_layouts[5])  # Blank layout
title1 = slide1.shapes.title
title1.text = "Test 1: Portfolio Allocation (PIE)"

chart1 = chart_builder.create_chart(
    slide1,
    allocation_data,
    left=Inches(1.0),
    top=Inches(2.0),
    width=Inches(8.0),
    height=Inches(4.0),
    title="Portfolio Allocation"
)

if chart1:
    print("✅ PIE chart created!")
    # Get position info
    for shape in slide1.shapes:
        if hasattr(shape, 'chart'):
            print(f"   Position: left={shape.left} EMUs ({shape.left/914400:.2f} inches)")
            print(f"   Expected: left={Inches(1.0)} EMUs (1.00 inches)")
            if abs(shape.left - Inches(1.0)) < 1000:
                print("   ✅ POSITION CORRECT!")
            else:
                print(f"   ❌ POSITION WRONG! Off by {abs(shape.left - Inches(1.0))} EMUs")
else:
    print("❌ PIE chart failed!")

# Test 2: LINE chart
print("\n📈 Test 2: Creating LINE chart from quarterly data...")
slide2 = prs.slides.add_slide(prs.slide_layouts[5])
title2 = slide2.shapes.title
title2.text = "Test 2: Quarterly Revenue (LINE)"

chart2 = chart_builder.create_chart(
    slide2,
    quarterly_data,
    left=Inches(1.0),
    top=Inches(2.0),
    width=Inches(8.0),
    height=Inches(4.0),
    title="Quarterly Trends"
)

if chart2:
    print("✅ LINE chart created!")
    for shape in slide2.shapes:
        if hasattr(shape, 'chart'):
            print(f"   Position: left={shape.left} EMUs ({shape.left/914400:.2f} inches)")
else:
    print("❌ LINE chart failed!")

# Test 3: COLUMN chart
print("\n📊 Test 3: Creating COLUMN chart from performance data...")
slide3 = prs.slides.add_slide(prs.slide_layouts[5])
title3 = slide3.shapes.title
title3.text = "Test 3: Top Performers (COLUMN)"

chart3 = chart_builder.create_chart(
    slide3,
    performance_data,
    left=Inches(1.0),
    top=Inches(2.0),
    width=Inches(8.0),
    height=Inches(4.0),
    title="Top Performers"
)

if chart3:
    print("✅ COLUMN chart created!")
    for shape in slide3.shapes:
        if hasattr(shape, 'chart'):
            print(f"   Position: left={shape.left} EMUs ({shape.left/914400:.2f} inches)")
else:
    print("❌ COLUMN chart failed!")

# Save
output_path = "simple_finance_charts_test.pptx"
prs.save(output_path)

print(f"\n✅ Test complete! Saved to: {output_path}")
print("📂 Open the file to verify charts are visible and positioned correctly")
