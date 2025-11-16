"""
Test Finance-Optimized Chart Detection
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.converter.enhanced_charts import EnhancedChartBuilder
import pandas as pd

print("\n" + "="*80)
print("🏦 TESTING FINANCE-OPTIMIZED CHART DETECTION")
print("="*80)

# Create chart builder
template_colors = {
    'chart_colors': [
        (123, 31, 162),   # Purple
        (171, 71, 188),
        (186, 104, 200),
        (206, 147, 216),
        (225, 190, 231)
    ]
}

builder = EnhancedChartBuilder(template_colors)

# Finance-specific test cases
finance_test_cases = [
    ("Portfolio Allocation", pd.DataFrame({
        'Sector': ['Technology', 'Financial Services', 'Healthcare', 'Energy', 'Consumer Goods'],
        'Allocation %': [30, 25, 20, 15, 10]
    }), 'PIE'),
    
    ("Sector Distribution", pd.DataFrame({
        'Sector': ['Tech', 'Finance', 'Healthcare'],
        'Weight': [0.40, 0.35, 0.25]
    }), 'PIE'),
    
    ("Quarterly Revenue", pd.DataFrame({
        'Quarter': ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024'],
        'Revenue': [1000000, 1200000, 1400000, 1600000]
    }), 'LINE'),
    
    ("YTD Performance", pd.DataFrame({
        'Month': ['Jan', 'Feb', 'Mar', 'Apr', 'May'],
        'Return %': [5.2, 6.1, 4.8, 7.3, 8.1]
    }), 'LINE'),
    
    ("Top 10 Stocks by Performance", pd.DataFrame({
        'Stock': ['AAPL', 'MSFT', 'GOOGL', 'AMZN', 'TSLA', 'META', 'NVDA', 'JPM', 'BAC', 'WFC'],
        'Return %': [45.2, 38.7, 32.1, 28.9, 65.4, 42.3, 71.8, 18.5, 22.1, 15.7]
    }), 'COLUMN'),
    
    ("Revenue by Product", pd.DataFrame({
        'Product': ['Product A', 'Product B', 'Product C', 'Product D'],
        'Sales': [500000, 450000, 380000, 320000]
    }), 'COLUMN'),
    
    ("Profit Comparison", pd.DataFrame({
        'Company': ['Company 1', 'Company 2', 'Company 3'],
        'Profit': [1200000, 950000, 880000]
    }), 'COLUMN'),
    
    ("Risk Breakdown", pd.DataFrame({
        'Risk Category': ['Low', 'Medium', 'High'],
        'Portfolio Share': [40, 45, 15]
    }), 'PIE'),
]

print("\n📊 TESTING FINANCE DATA SCENARIOS:")
print("-" * 80)

correct = 0
total = len(finance_test_cases)

for name, df, expected_type in finance_test_cases:
    detected = builder.detect_chart_type(df).upper()
    is_correct = detected == expected_type
    correct += is_correct
    
    status = "✅" if is_correct else "❌"
    print(f"\n{status} {name}")
    print(f"   Columns: {list(df.columns)}")
    print(f"   Expected: {expected_type}")
    print(f"   Detected: {detected}")
    if not is_correct:
        print(f"   ⚠️ MISMATCH! Should be {expected_type}")

print("\n" + "="*80)
print(f"📈 RESULTS: {correct}/{total} correct ({(correct/total)*100:.1f}%)")
print("="*80)

print("\n🏦 FINANCE CHART RULES:")
print("-" * 80)
print("""
1. PIE CHARTS for:
   - Portfolio allocation
   - Sector distribution
   - Any data with "allocation", "portfolio", "sector", "%"
   
2. LINE CHARTS for:
   - Quarterly/monthly trends
   - Performance over time
   - Any data with "quarter", "month", "date", "trend", "ytd"
   
3. COLUMN CHARTS (Vertical Bars) for:
   - Top performers/rankings
   - Revenue/sales/profit comparisons
   - Any data with "top", "performance", "revenue", "sales", "profit"
   
4. DEFAULT: COLUMN charts (finance standard for comparisons)
""")

print("\n💡 BENEFITS:")
print("-" * 80)
print("""
✅ PIE charts show percentages + values (e.g., "30M (25%)")
✅ COLUMN charts show data labels on top of bars
✅ LINE charts have markers for clarity
✅ All charts use template colors (purple from royal_purple)
✅ Finance-appropriate formatting and styling
""")

print("\n" + "="*80)
