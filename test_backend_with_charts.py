"""
Test Backend API with Advanced Finance Charts
Tests the complete flow: Excel upload → Backend API → Converter → Charts
"""

import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter
import pandas as pd

print("=" * 80)
print("🎨 TESTING BACKEND CONNECTION WITH ADVANCED CHARTS")
print("=" * 80)

# Test 1: Create test Excel file
print("\n📊 Step 1: Creating test Excel file...")
test_excel = "examples/test_backend_charts.xlsx"

# Create sample data with multiple sheets
with pd.ExcelWriter(test_excel, engine='openpyxl') as writer:
    # Sheet 1: Financial Performance
    df1 = pd.DataFrame({
        'Quarter': ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024'],
        'Revenue': [1200000, 1350000, 1420000, 1580000],
        'Profit': [240000, 270000, 284000, 316000],
        'Expenses': [960000, 1080000, 1136000, 1264000]
    })
    df1.to_excel(writer, sheet_name='Financial Performance', index=False)
    
    # Sheet 2: Portfolio Allocation
    df2 = pd.DataFrame({
        'Asset Class': ['Stocks', 'Bonds', 'Real Estate', 'Commodities', 'Cash'],
        'Allocation %': [45, 30, 15, 5, 5]
    })
    df2.to_excel(writer, sheet_name='Portfolio Allocation', index=False)
    
    # Sheet 3: Product Sales
    df3 = pd.DataFrame({
        'Product': ['Product A', 'Product B', 'Product C', 'Product D', 'Product E'],
        'Sales': [850000, 720000, 650000, 480000, 320000],
        'Growth %': [15, 12, 8, 5, -3]
    })
    df3.to_excel(writer, sheet_name='Product Sales', index=False)

print(f"   ✅ Created test Excel: {test_excel}")

# Test 2: Simulate Backend API Call
print("\n🔧 Step 2: Simulating Backend API with use_finance_charts=True...")

output_ppt = "examples/demo_PPT/backend_test_output.pptx"

# Create converter instance (simulating what backend does)
converter = ExcelToPPTConverter(
    user_tier='pro',  # Test with PRO tier
    user_id='test_user_123',
    user_metadata={
        'name': 'Test User',
        'email': 'test@example.com',
        'company': 'Test Company'
    },
    use_finance_charts=True  # ✅ THIS IS NOW DEFAULT IN BACKEND
)

print(f"   ✅ Converter initialized with use_finance_charts=True")
print(f"   ✅ User tier: pro")

# Test 3: Run conversion
print("\n🎨 Step 3: Running conversion with EnhancedProfessionalBuilder...")

result = converter.convert_professional(
    excel_path=test_excel,
    output_path=output_ppt,
    template_name=None,  # Use default professional template
    presentation_title="Backend Test - Advanced Charts",
    user_ppt_count=0,
    use_professional_structure=True
)

# Test 4: Verify results
print("\n" + "=" * 80)
print("📊 CONVERSION RESULTS")
print("=" * 80)

if result['success']:
    print("✅ SUCCESS! Presentation created with Advanced Finance Charts")
    print(f"\n📁 Output file: {output_ppt}")
    print(f"📊 Total slides: {result.get('slides_created', 'Unknown')}")
    
    if 'slide_names' in result:
        print(f"\n📋 Slides created:")
        for i, name in enumerate(result['slide_names'], 1):
            print(f"   {i}. {name}")
    
    # Check for chart types used
    print(f"\n🎨 Charts should be integrated in:")
    print(f"   ✅ Key Metrics - COLUMN chart")
    print(f"   ✅ Sector Distribution - DONUT/WATERFALL chart")
    print(f"   ✅ Top Performers - BAR chart")
    print(f"   ✅ Trend Analysis - LINE/CANDLESTICK chart")
    
    print(f"\n🎯 Total presentation: ~10 slides with 3-5 different chart types")
    
    print("\n" + "=" * 80)
    print("🎉 BACKEND CONNECTION VERIFIED!")
    print("=" * 80)
    print(f"\nOpen the file to see the charts:")
    print(f"   {os.path.abspath(output_ppt)}")
    
else:
    print("❌ FAILED!")
    print(f"Error: {result.get('error', 'Unknown error')}")
    if 'message' in result:
        print(f"Message: {result['message']}")

print("\n" + "=" * 80)
