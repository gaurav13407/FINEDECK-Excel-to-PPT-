"""
Test script to verify enhanced chart detection at runtime
This simulates what actually happens when the backend processes an Excel file
"""

import pandas as pd
import numpy as np
import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from converter.enhanced_charts import EnhancedChartBuilder

print("=" * 80)
print("RUNTIME CHART DETECTION DIAGNOSTIC")
print("=" * 80)

# Test 1: Load actual Excel file that user uploaded
excel_file = r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\examples\Portfolio Allocation Data.xlsx"

if os.path.exists(excel_file):
    print(f"\n📂 Loading: {excel_file}")
    df = pd.read_excel(excel_file)
    print(f"   Columns: {df.columns.tolist()}")
    print(f"   Shape: {df.shape}")
    print(f"   First few rows:")
    print(df.head())
    
    # Initialize enhanced chart builder
    template_colors = {
        'navy': (31, 73, 125),
        'light_blue': (68, 114, 196),
        'teal': (91, 155, 213),
        'gray': (165, 165, 165),
        'dark_gray': (89, 89, 89),
        'accent_orange': (237, 125, 49),
        'accent_green': (112, 173, 71)
    }
    
    builder = EnhancedChartBuilder(template_colors)
    
    # Test detection
    detected_type = builder.detect_chart_type(df)
    print(f"\n🔍 DETECTION RESULT:")
    print(f"   Chart Type: {detected_type.upper()}")
    print(f"   Expected: PIE (for portfolio allocation)")
    
    if detected_type == 'pie':
        print("   ✅ CORRECT - Finance detection working!")
    else:
        print(f"   ❌ WRONG - Expected PIE but got {detected_type.upper()}")
        
        # Debug why it failed
        col_names_lower = ' '.join([str(col).lower() for col in df.columns])
        print(f"\n🐛 DEBUG INFO:")
        print(f"   Column names (lowercase): {col_names_lower}")
        
        # Check each detection category
        line_keywords = ['date', 'time', 'quarter', 'month', 'year', 'q1', 'q2', 'q3', 'q4',
                        'trend', 'growth', 'ytd', 'mtd']
        pie_keywords = ['allocation', 'portfolio', 'sector', 'distribution', 'breakdown',
                       'composition', 'mix', 'share', 'weight']
        
        print(f"   Matches LINE keywords: {any(kw in col_names_lower for kw in line_keywords)}")
        print(f"   Matches PIE keywords: {any(kw in col_names_lower for kw in pie_keywords)}")
        
        if any(kw in col_names_lower for kw in line_keywords):
            matching = [kw for kw in line_keywords if kw in col_names_lower]
            print(f"      Matching LINE keywords: {matching}")
        if any(kw in col_names_lower for kw in pie_keywords):
            matching = [kw for kw in pie_keywords if kw in col_names_lower]
            print(f"      Matching PIE keywords: {matching}")

else:
    print(f"❌ File not found: {excel_file}")

print("\n" + "=" * 80)

# Test 2: Check what data structure sector distribution receives
print("\nTEST 2: SECTOR DISTRIBUTION DATA STRUCTURE")
print("=" * 80)

# Simulate what _create_sector_distribution does
if os.path.exists(excel_file):
    df = pd.read_excel(excel_file)
    
    # Find category column (what the backend does)
    text_cols = df.select_dtypes(include=['object']).columns
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    
    print(f"Text columns: {text_cols.tolist()}")
    print(f"Numeric columns: {numeric_cols.tolist()}")
    
    if len(text_cols) > 0 and len(numeric_cols) > 0:
        category_col = text_cols[0]
        value_col = numeric_cols[0]
        
        print(f"\nUsing: category={category_col}, value={value_col}")
        
        # Aggregate by category (what backend does)
        sector_data = df.groupby(category_col)[value_col].sum().sort_values(ascending=False).head(8)
        
        # Convert to DataFrame (what backend does)
        sector_df = sector_data.reset_index()
        sector_df.columns = [category_col, value_col]
        
        print(f"\nSector DataFrame passed to auto_create_chart:")
        print(sector_df)
        print(f"Columns: {sector_df.columns.tolist()}")
        
        # Test detection on THIS data structure
        detected_type = builder.detect_chart_type(sector_df)
        print(f"\n🔍 DETECTION ON SECTOR DATA:")
        print(f"   Chart Type: {detected_type.upper()}")
        print(f"   Expected: PIE")
        
        if detected_type == 'pie':
            print("   ✅ CORRECT")
        else:
            print(f"   ❌ WRONG - Got {detected_type.upper()}")
            
            # Debug
            col_names_lower = ' '.join([str(col).lower() for col in sector_df.columns])
            print(f"   Column names: {col_names_lower}")

print("\n" + "=" * 80)

# Test 3: Check actual financial data files
print("\nTEST 3: ACTUAL USER FILES (AMZN, TSLA, GOOGL)")
print("=" * 80)

test_files = [
    r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\examples\Sample_pnl.xlsx",
    r"c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)\examples\Risk Metrics Data.xlsx",
]

for test_file in test_files:
    if os.path.exists(test_file):
        print(f"\n📂 {os.path.basename(test_file)}")
        try:
            df = pd.read_excel(test_file)
            print(f"   Columns: {df.columns.tolist()}")
            
            detected_type = builder.detect_chart_type(df)
            print(f"   Detected: {detected_type.upper()}")
            
            # Show what keywords matched
            col_names_lower = ' '.join([str(col).lower() for col in df.columns])
            
            if 'date' in col_names_lower or 'quarter' in col_names_lower or 'month' in col_names_lower:
                print("   → Contains TIME keywords (should be LINE)")
            elif 'allocation' in col_names_lower or 'portfolio' in col_names_lower or 'sector' in col_names_lower:
                print("   → Contains ALLOCATION keywords (should be PIE)")
            elif 'top' in col_names_lower or 'rank' in col_names_lower or 'performance' in col_names_lower:
                print("   → Contains PERFORMANCE keywords (should be COLUMN)")
            
        except Exception as e:
            print(f"   ❌ Error: {e}")

print("\n" + "=" * 80)
print("RECOMMENDATION:")
print("=" * 80)
print("""
If the detection is CORRECT in this test but WRONG in actual PPTs:
1. The enhanced_chart_builder IS initialized correctly ✅
2. The finance detection logic IS working ✅
3. BUT: The actual data being passed to auto_create_chart() might be different

Check:
- Backend terminal logs for "✨ Using EnhancedChartBuilder..." messages
- What columns the actual Excel file has vs. what reaches auto_create_chart()
- If there's data transformation happening before chart creation
""")
