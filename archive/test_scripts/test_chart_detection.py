"""
Test chart detection on YOUR actual Excel file
"""
import sys
import pandas as pd
from src.converter.simple_finance_charts import SimpleFinanceChartBuilder

if len(sys.argv) < 2:
    print("Usage: python test_chart_detection.py <path_to_excel_file>")
    print("\nExample:")
    print('  python test_chart_detection.py "examples/Portfolio Allocation Data.xlsx"')
    sys.exit(1)

excel_file = sys.argv[1]

print("="*80)
print(f"TESTING CHART DETECTION ON: {excel_file}")
print("="*80)

try:
    # Read all sheets
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    
    print(f"\n📊 Found {len(excel_data)} sheet(s)")
    
    builder = SimpleFinanceChartBuilder()
    
    for sheet_name, df in excel_data.items():
        print(f"\n{'='*80}")
        print(f"SHEET: {sheet_name}")
        print(f"{'='*80}")
        
        print(f"Rows: {len(df)}")
        print(f"Columns: {list(df.columns)}")
        
        # Get column types
        text_cols = df.select_dtypes(include=['object']).columns.tolist()
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        
        print(f"\nText columns: {text_cols}")
        print(f"Numeric columns: {numeric_cols}")
        
        if len(numeric_cols) == 0:
            print("\n⚠️  No numeric columns - CANNOT create chart")
            continue
        
        # Test detection
        print(f"\n🔍 Testing chart type detection...")
        chart_type = builder.detect_chart_type(df)
        
        print(f"\n✅ DETECTED CHART TYPE: {chart_type.upper()}")
        
        # Show what triggered it
        col_names_lower = ' '.join([str(col).lower() for col in df.columns])
        print(f"\nColumn names (lowercase): {col_names_lower}")
        
        print("\n📝 Detection Logic:")
        
        # Check time-series
        time_keywords = ['date', 'quarter', 'month', 'year', 'q1', 'q2', 'q3', 'q4', 'ytd', 'period']
        time_matches = [kw for kw in time_keywords if kw in col_names_lower]
        if time_matches:
            print(f"  ⏰ Time-series keywords found: {time_matches} → Would create LINE chart")
        
        # Check allocation
        allocation_keywords = ['allocation', 'portfolio', 'sector', 'distribution', 'breakdown', 'composition']
        alloc_matches = [kw for kw in allocation_keywords if kw in col_names_lower]
        if alloc_matches:
            print(f"  🥧 Allocation keywords found: {alloc_matches} → Would create PIE chart")
        
        # Check performance
        performance_keywords = ['top', 'rank', 'performance', 'revenue', 'sales', 'profit', 'growth', 'comparison']
        perf_matches = [kw for kw in performance_keywords if kw in col_names_lower]
        if perf_matches:
            print(f"  📊 Performance keywords found: {perf_matches} → Would create COLUMN chart")
        
        # Show sample data
        print(f"\n📋 Sample Data (first 3 rows):")
        print(df.head(3).to_string())
        
        # Recommendation
        print(f"\n💡 RECOMMENDATION:")
        if chart_type == 'pie':
            print(f"  This data looks like allocation/distribution")
            print(f"  → PIE chart is appropriate ✅")
        elif chart_type == 'line':
            print(f"  This data looks like time-series")
            print(f"  → LINE chart is appropriate ✅")
        elif chart_type == 'column':
            print(f"  This data looks like comparison/performance")
            print(f"  → COLUMN chart is appropriate ✅")
        
        print(f"\n🎯 TO CHANGE CHART TYPE:")
        print(f"  Rename columns to include keywords:")
        print(f"  - For PIE: Include 'allocation', 'portfolio', 'sector', 'distribution'")
        print(f"  - For LINE: Include 'date', 'quarter', 'month', 'year', 'Q1', 'Q2', etc.")
        print(f"  - For COLUMN: Include 'performance', 'revenue', 'sales', 'top', 'rank'")

except FileNotFoundError:
    print(f"\n❌ ERROR: File not found: {excel_file}")
except Exception as e:
    print(f"\n❌ ERROR: {e}")
    import traceback
    traceback.print_exc()

print(f"\n{'='*80}")
