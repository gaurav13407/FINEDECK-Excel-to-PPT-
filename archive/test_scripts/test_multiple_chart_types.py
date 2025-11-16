"""
Test creating multiple chart types from a single Excel file
Demonstrates: LINE, AREA, COLUMN, BAR, DONUT, STACKED_COLUMN, WATERFALL, CANDLESTICK, SCATTER charts
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys
import os

# Add project to path
sys.path.insert(0, os.path.abspath('.'))

from src.converter.excel_to_ppt_converter import ExcelToPPTConverter

def create_comprehensive_test_excel():
    """Create an Excel file with diverse data that will trigger multiple chart types"""
    
    print("\n📊 Creating comprehensive test Excel file...")
    
    # Create writer
    excel_path = "examples/comprehensive_financial_data.xlsx"
    writer = pd.ExcelWriter(excel_path, engine='openpyxl')
    
    # Sheet 1: Quarterly Performance (for LINE, AREA, COLUMN charts)
    quarters = ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024', 'Q1 2025', 'Q2 2025']
    quarterly_data = pd.DataFrame({
        'Quarter': quarters,
        'Revenue': [125000, 142000, 158000, 189000, 215000, 248000],
        'Profit': [35000, 42000, 48000, 62000, 78000, 92000],
        'Expenses': [90000, 100000, 110000, 127000, 137000, 156000],
        'Operating Income': [45000, 52000, 61000, 71000, 88000, 102000]
    })
    quarterly_data.to_excel(writer, sheet_name='Quarterly Performance', index=False)
    print("   ✅ Quarterly Performance sheet (LINE, AREA, COLUMN)")
    
    # Sheet 2: Department Performance (for BAR chart)
    departments = ['Sales', 'Marketing', 'Product', 'Engineering', 'Customer Success', 
                   'Operations', 'Finance', 'HR', 'R&D', 'IT', 'Legal', 'Admin']
    dept_data = pd.DataFrame({
        'Department': departments,
        'Revenue Generated': [2500000, 450000, 680000, 320000, 890000, 240000, 
                             180000, 120000, 550000, 280000, 95000, 75000],
        'Team Size': [45, 12, 28, 65, 22, 18, 8, 6, 35, 15, 4, 8]
    })
    dept_data.to_excel(writer, sheet_name='Department Rankings', index=False)
    print("   ✅ Department Rankings sheet (BAR chart)")
    
    # Sheet 3: Portfolio Allocation (for DONUT/PIE chart)
    allocation_data = pd.DataFrame({
        'Asset Class': ['Stocks', 'Bonds', 'Real Estate', 'Commodities', 'Cash'],
        'Allocation %': [45, 25, 15, 10, 5],
        'Value': [4500000, 2500000, 1500000, 1000000, 500000]
    })
    allocation_data.to_excel(writer, sheet_name='Portfolio Allocation', index=False)
    print("   ✅ Portfolio Allocation sheet (DONUT chart)")
    
    # Sheet 4: Monthly Composition (for STACKED_COLUMN)
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
    composition_data = pd.DataFrame({
        'Month': months,
        'Product A': [120000, 135000, 145000, 152000, 168000, 185000],
        'Product B': [85000, 92000, 98000, 105000, 112000, 125000],
        'Product C': [65000, 68000, 72000, 78000, 85000, 95000],
        'Product D': [45000, 48000, 52000, 55000, 62000, 70000]
    })
    composition_data.to_excel(writer, sheet_name='Product Mix', index=False)
    print("   ✅ Product Mix sheet (STACKED_COLUMN)")
    
    # Sheet 5: P&L Waterfall Data (for WATERFALL chart)
    pl_data = pd.DataFrame({
        'Category': ['Revenue', 'COGS', 'Gross Profit', 'R&D', 'Sales & Marketing', 
                    'G&A', 'Operating Income', 'Interest', 'Taxes', 'Net Income'],
        'Amount': [10000000, -4500000, 5500000, -800000, -1200000, 
                  -600000, 2900000, -150000, -825000, 1925000]
    })
    pl_data.to_excel(writer, sheet_name='P&L Waterfall', index=False)
    print("   ✅ P&L Waterfall sheet (WATERFALL chart)")
    
    # Sheet 6: Stock Prices (for CANDLESTICK chart)
    dates = pd.date_range('2025-01-01', periods=20, freq='D')
    base_price = 150
    stock_data = pd.DataFrame({
        'Date': [d.strftime('%Y-%m-%d') for d in dates],
        'Open': [base_price + np.random.uniform(-5, 5) + i*0.5 for i in range(20)],
        'High': [base_price + np.random.uniform(2, 8) + i*0.5 for i in range(20)],
        'Low': [base_price + np.random.uniform(-8, -2) + i*0.5 for i in range(20)],
        'Close': [base_price + np.random.uniform(-4, 6) + i*0.5 for i in range(20)],
        'Volume': [np.random.randint(2000000, 4000000) for _ in range(20)]
    })
    stock_data.to_excel(writer, sheet_name='Stock Prices', index=False)
    print("   ✅ Stock Prices sheet (CANDLESTICK chart)")
    
    # Sheet 7: Risk-Return Analysis (for SCATTER chart)
    investments = ['Bonds', 'Index Fund', 'Blue Chip', 'Growth Stocks', 'Small Cap', 
                   'Emerging Markets', 'Tech Stocks', 'Crypto']
    scatter_data = pd.DataFrame({
        'Investment': investments,
        'Risk (Volatility %)': [5.2, 8.5, 12.3, 15.8, 18.2, 22.5, 25.8, 35.5],
        'Return (Annual %)': [4.5, 7.2, 9.8, 12.5, 14.8, 16.2, 18.5, 22.8],
        'Market Cap (B)': [500, 350, 280, 180, 120, 95, 85, 45]
    })
    scatter_data.to_excel(writer, sheet_name='Risk-Return', index=False)
    print("   ✅ Risk-Return sheet (SCATTER chart)")
    
    writer.close()
    print(f"\n✅ Excel file created: {excel_path}")
    print(f"   📊 7 sheets with diverse data types")
    print(f"   🎯 Will generate: LINE, AREA, COLUMN, BAR, DONUT, STACKED_COLUMN, WATERFALL, CANDLESTICK, SCATTER\n")
    
    return excel_path

def test_multiple_charts():
    """Test converting Excel to PPT with multiple chart types"""
    
    print("\n" + "="*80)
    print("🎨 TESTING MULTIPLE CHART TYPES FROM SINGLE EXCEL FILE")
    print("="*80 + "\n")
    
    # Create test Excel file
    excel_path = create_comprehensive_test_excel()
    
    # Convert to PowerPoint
    print("🔄 Converting Excel to PowerPoint...")
    print("   Using: EnhancedProfessionalBuilder with AdvancedFinanceChartBuilder")
    print("   Expected: Multiple slides with different chart types\n")
    
    try:
        converter = ExcelToPPTConverter(
            user_tier='pro',  # Use pro tier for full features
            use_finance_charts=True
        )
        
        output_path = "examples/demo_PPT/multiple_charts_showcase.pptx"
        
        result = converter.convert(
            excel_path=excel_path,
            output_path=output_path,
            template_name="Modern Corporate",
            presentation_title="Comprehensive Financial Analysis"
        )
        
        if result.get('success'):
            print("\n" + "="*80)
            print("🎉 SUCCESS! Multiple chart types created!")
            print("="*80)
            print(f"📁 Output: {output_path}")
            print(f"📊 Slides created: {result.get('slides_created', 'N/A')}")
            print(f"📋 Slide names: {result.get('slide_names', [])}")
            print("\n✅ Expected chart types in presentation:")
            print("   1️⃣  LINE Chart (Revenue trends)")
            print("   2️⃣  AREA Chart (Cumulative growth)")
            print("   3️⃣  COLUMN Chart (Performance comparison)")
            print("   4️⃣  BAR Chart (Department rankings)")
            print("   5️⃣  DONUT Chart (Portfolio allocation)")
            print("   6️⃣  STACKED COLUMN (Product composition)")
            print("   7️⃣  WATERFALL Chart (P&L flow)")
            print("   8️⃣  CANDLESTICK Chart (Stock prices)")
            print("   9️⃣  SCATTER Chart (Risk-return analysis)")
            print("="*80 + "\n")
        else:
            print(f"\n❌ Conversion failed: {result.get('error', 'Unknown error')}")
            
    except Exception as e:
        print(f"\n❌ Error during conversion: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_multiple_charts()
