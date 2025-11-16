"""
Test Suite for Advanced Finance Chart Builder
Demonstrates all 12+ chart types with sample financial data
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from pptx import Presentation
from pptx.util import Inches
from src.converter.advanced_finance_charts import AdvancedFinanceChartBuilder

def create_sample_presentations():
    """Create sample presentations showcasing all chart types"""
    
    print("\n" + "="*80)
    print("🎨 ADVANCED FINANCE CHART BUILDER - COMPREHENSIVE TEST")
    print("="*80)
    print("Creating sample presentations with all 12+ chart types...")
    print("="*80 + "\n")
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # ========================================================================
    # 1️⃣ PERFORMANCE & GROWTH TRACKING
    # ========================================================================
    
    print("\n📈 1️⃣ TESTING PERFORMANCE & GROWTH CHARTS")
    print("-" * 80)
    
    # Test 1: Line Chart - Revenue Growth
    print("\n🔹 Test 1: LINE CHART (Revenue Growth)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    
    revenue_data = pd.DataFrame({
        'Quarter': ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024', 'Q1 2025'],
        'Revenue': [125000, 142000, 158000, 189000, 215000],
        'Profit': [35000, 42000, 48000, 62000, 78000],
        'Expenses': [90000, 100000, 110000, 127000, 137000]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(revenue_data, chart_type='LINE', title='Quarterly Revenue Growth')
    print("✅ Line chart created")
    
    # Test 2: Area Chart - Cumulative Profit
    print("\n🔹 Test 2: AREA CHART (Cumulative Profit)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    cumulative_data = pd.DataFrame({
        'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        'Net Income': [45000, 52000, 61000, 58000, 67000, 75000],
        'Operating Profit': [38000, 44000, 51000, 49000, 56000, 63000]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(cumulative_data, chart_type='AREA', title='Cumulative Profit Trend')
    print("✅ Area chart created")
    
    # Test 3: Column Chart - YoY Comparison
    print("\n🔹 Test 3: COLUMN CHART (YoY Comparison)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    yoy_data = pd.DataFrame({
        'Region': ['North America', 'Europe', 'Asia Pacific', 'Latin America'],
        '2023': [450000, 380000, 520000, 180000],
        '2024': [520000, 425000, 680000, 225000]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(yoy_data, chart_type='COLUMN', title='YoY Regional Performance')
    print("✅ Column chart created")
    
    # ========================================================================
    # 2️⃣ PORTFOLIO & ASSET ANALYSIS
    # ========================================================================
    
    print("\n\n💰 2️⃣ TESTING PORTFOLIO & ASSET CHARTS")
    print("-" * 80)
    
    # Test 4: Pie Chart - Asset Allocation
    print("\n🔹 Test 4: PIE CHART (Asset Allocation)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    allocation_data = pd.DataFrame({
        'Asset Class': ['Stocks', 'Bonds', 'Real Estate', 'Commodities', 'Cash'],
        'Allocation': [45, 25, 15, 10, 5]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(allocation_data, chart_type='PIE', title='Portfolio Asset Allocation')
    print("✅ Pie chart created")
    
    # Test 5: Donut Chart - Sector Distribution
    print("\n🔹 Test 5: DONUT CHART (Sector Distribution)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    sector_data = pd.DataFrame({
        'Sector': ['Technology', 'Healthcare', 'Finance', 'Energy', 'Consumer', 'Industrial'],
        'Portfolio Weight': [28, 18, 16, 12, 15, 11]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(sector_data, chart_type='DONUT', title='Sector Distribution')
    print("✅ Donut chart created")
    
    # Test 6: Stacked Column - Portfolio Composition Over Time
    print("\n🔹 Test 6: STACKED COLUMN (Portfolio Composition)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    composition_data = pd.DataFrame({
        'Quarter': ['Q1', 'Q2', 'Q3', 'Q4'],
        'Stocks': [450000, 480000, 520000, 560000],
        'Bonds': [250000, 260000, 270000, 280000],
        'Real Estate': [150000, 155000, 160000, 165000],
        'Cash': [50000, 45000, 40000, 35000]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(composition_data, chart_type='STACKED_COLUMN', 
                        title='Portfolio Composition Over Time')
    print("✅ Stacked column chart created")
    
    # ========================================================================
    # 3️⃣ P&L & CASH FLOW
    # ========================================================================
    
    print("\n\n📊 3️⃣ TESTING P&L & CASH FLOW CHARTS")
    print("-" * 80)
    
    # Test 7: Waterfall Chart - P&L Bridge
    print("\n🔹 Test 7: WATERFALL CHART (P&L Bridge)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    waterfall_data = pd.DataFrame({
        'Category': ['Revenue', 'COGS', 'Gross Profit', 'OpEx', 'EBITDA', 'D&A', 'EBIT', 'Tax', 'Net Income'],
        'Amount': [1000000, -450000, 550000, -280000, 270000, -35000, 235000, -70500, 164500]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(waterfall_data, chart_type='WATERFALL', title='P&L Waterfall: Revenue to Net Income')
    print("✅ Waterfall chart created")
    
    # Test 8: Stacked Bar - Expense Breakdown
    print("\n🔹 Test 8: STACKED BAR (Expense Breakdown)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    expense_data = pd.DataFrame({
        'Quarter': ['Q1', 'Q2', 'Q3', 'Q4'],
        'Salaries': [180000, 185000, 190000, 195000],
        'Marketing': [45000, 52000, 48000, 55000],
        'R&D': [65000, 68000, 72000, 75000],
        'Operations': [38000, 40000, 42000, 44000]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(expense_data, chart_type='STACKED_BAR', title='Quarterly Expense Composition')
    print("✅ Stacked bar chart created")
    
    # Test 9: Bar Chart - Department Performance
    print("\n🔹 Test 9: BAR CHART (Department Performance)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    dept_data = pd.DataFrame({
        'Department': ['Sales', 'Marketing', 'Product', 'Engineering', 'Customer Success', 
                      'Operations', 'Finance', 'HR'],
        'Revenue Generated': [2500000, 450000, 680000, 320000, 890000, 240000, 180000, 120000]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(dept_data, chart_type='BAR', title='Revenue by Department')
    print("✅ Bar chart created")
    
    # ========================================================================
    # 4️⃣ MARKET & STOCK ANALYSIS
    # ========================================================================
    
    print("\n\n📈 4️⃣ TESTING MARKET & STOCK ANALYSIS CHARTS")
    print("-" * 80)
    
    # Test 10: Candlestick Chart - Stock Price Movement
    print("\n🔹 Test 10: CANDLESTICK CHART (Stock Prices)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Generate realistic stock data
    dates = pd.date_range('2025-01-01', periods=10, freq='D')
    stock_data = pd.DataFrame({
        'Date': [d.strftime('%m/%d') for d in dates],
        'Open': [150.2, 152.5, 151.8, 153.2, 154.5, 153.8, 155.2, 156.8, 155.5, 157.2],
        'High': [153.8, 154.2, 154.5, 155.8, 156.2, 156.5, 157.8, 158.2, 157.8, 159.5],
        'Low': [149.5, 151.2, 150.8, 152.5, 153.2, 152.8, 154.5, 155.2, 154.8, 156.2],
        'Close': [152.5, 151.8, 153.2, 154.5, 153.8, 155.2, 156.8, 155.5, 157.2, 158.8],
        'Volume': [2500000, 2800000, 2300000, 3100000, 2700000, 2900000, 3200000, 2600000, 2800000, 3000000]
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(stock_data, chart_type='CANDLESTICK', title='Stock Price Movement (OHLC)')
    print("✅ Candlestick chart created")
    
    # Test 11: Scatter Plot - Risk vs Return
    print("\n🔹 Test 11: SCATTER CHART (Risk-Return Analysis)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    risk_return_data = pd.DataFrame({
        'Risk (Volatility %)': [5.2, 8.5, 12.3, 15.8, 18.2, 22.5, 25.8, 28.3],
        'Return (Annual %)': [4.5, 7.2, 9.8, 12.5, 14.8, 16.2, 18.5, 19.8],
        'Asset': ['Bonds', 'Balanced Fund', 'Blue Chip', 'Growth', 'Small Cap', 
                 'Emerging Markets', 'Tech Stocks', 'Crypto']
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(risk_return_data, chart_type='SCATTER', title='Risk-Return Profile')
    print("✅ Scatter chart created")
    
    # Test 12: Bubble Chart - Multi-Dimensional Analysis
    print("\n🔹 Test 12: BUBBLE CHART (Market Cap, Risk, Return)")
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    bubble_data = pd.DataFrame({
        'Market Cap (Billion)': [500, 350, 280, 180, 120, 95],
        'Annual Return (%)': [15.2, 18.5, 22.3, 25.8, 28.5, 32.1],
        'Risk Score': [3.5, 4.8, 6.2, 7.5, 8.8, 9.5],
        'Company': ['AAPL', 'MSFT', 'GOOGL', 'AMZN', 'TSLA', 'NVDA']
    })
    
    builder = AdvancedFinanceChartBuilder(slide, position=(1, 1), size=(8, 5))
    builder.create_chart(bubble_data, chart_type='BUBBLE', 
                        title='Stock Analysis: Market Cap vs Return vs Risk')
    print("✅ Bubble chart created")
    
    # ========================================================================
    # SAVE PRESENTATION
    # ========================================================================
    
    output_path = "examples/demo_PPT/advanced_finance_charts_showcase.pptx"
    prs.save(output_path)
    
    print("\n" + "="*80)
    print("🎉 SUCCESS! All 12 chart types created successfully!")
    print("="*80)
    print(f"📁 Saved to: {output_path}")
    print("\n📊 Chart Types Demonstrated:")
    print("   1️⃣  LINE Chart (Revenue Growth)")
    print("   2️⃣  AREA Chart (Cumulative Profit)")
    print("   3️⃣  COLUMN Chart (YoY Comparison)")
    print("   4️⃣  PIE Chart (Asset Allocation)")
    print("   5️⃣  DONUT Chart (Sector Distribution)")
    print("   6️⃣  STACKED COLUMN Chart (Portfolio Composition)")
    print("   7️⃣  WATERFALL Chart (P&L Bridge)")
    print("   8️⃣  STACKED BAR Chart (Expense Breakdown)")
    print("   9️⃣  BAR Chart (Department Performance)")
    print("   🔟 CANDLESTICK Chart (Stock Prices)")
    print("   1️⃣1️⃣ SCATTER Chart (Risk-Return)")
    print("   1️⃣2️⃣ BUBBLE Chart (Multi-Dimensional)")
    print("="*80 + "\n")

if __name__ == "__main__":
    create_sample_presentations()
