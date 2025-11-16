"""
Test excel_reader with actual sample files to verify column cleaning
"""

import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from converter.excel_reader import excel_reader, get_sheet_names, excel_reader_all_sheets

def test_sample_files():
    """Test column cleaning with actual sample files."""
    
    sample_files = [
        "examples/Portfolio Allocation Data.xlsx",
        "examples/Risk Metrics Data.xlsx",
        "examples/Sample_pnl.xlsx",
        "examples/Scenario Comparison (Bull_Bear_Base).xlsx"
    ]
    
    for file_path in sample_files:
        if not os.path.exists(file_path):
            print(f"⚠️  File not found: {file_path}")
            continue
        
        print("\n" + "="*70)
        print(f"📁 FILE: {os.path.basename(file_path)}")
        print("="*70)
        
        try:
            # Get sheet names
            sheets = get_sheet_names(file_path)
            print(f"\n📋 Sheets: {sheets}")
            
            # Read all sheets
            all_data = excel_reader_all_sheets(file_path)
            
            for sheet_name, df in all_data.items():
                print(f"\n📊 Sheet: '{sheet_name}'")
                print("-" * 70)
                
                if df is not None and not df.empty:
                    print(f"   Shape: {df.shape[0]} rows × {df.shape[1]} columns")
                    print(f"   Columns: {df.columns.tolist()}")
                    
                    # Check for cleaning issues
                    issues = []
                    for col in df.columns:
                        col_str = str(col)
                        if 'unnamed' in col_str.lower():
                            issues.append(f"❌ Unnamed: '{col}'")
                        elif isinstance(col, str):
                            # Check for numeric prefix like "1. Column"
                            import re
                            if re.match(r'^\d+\.', col):
                                issues.append(f"❌ Numeric prefix: '{col}'")
                            # Check for leading/trailing whitespace
                            if col != col.strip():
                                issues.append(f"❌ Whitespace: '{col}'")
                    
                    if issues:
                        print(f"   ⚠️  ISSUES:")
                        for issue in issues:
                            print(f"      {issue}")
                    else:
                        print(f"   ✅ All columns clean!")
                    
                    # Show sample data
                    print(f"\n   Sample data (first 3 rows):")
                    print("   " + "-" * 66)
                    sample = df.head(3)
                    for idx, row in sample.iterrows():
                        print(f"   Row {idx+1}: {dict(row)}")
                    
                else:
                    print(f"   ✗ Empty or failed to read")
        
        except Exception as e:
            print(f"   ❌ ERROR: {str(e)}")
    
    print("\n" + "="*70)
    print("✨ TESTING COMPLETE")
    print("="*70)


if __name__ == "__main__":
    test_sample_files()
