"""
Test script to create a multi-sheet Excel file with messy columns
and test the excel_reader module's cleaning capabilities.
"""

import pandas as pd
import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from converter.excel_reader import excel_reader, get_sheet_names, excel_reader_all_sheets

def create_test_excel():
    """Create a test Excel file with messy columns across multiple sheets."""
    
    # Sheet 1: Portfolio Data with unnamed columns and numeric prefixes
    sheet1_data = {
        'Unnamed: 0': ['', '', 'Stock A', 'Stock B', 'Stock C'],
        '1. Asset Name': ['', '', 'Apple Inc.', 'Microsoft Corp.', 'Google LLC'],
        '2. Sector': ['', '', 'Technology', 'Technology', 'Technology'],
        'Value': ['', '', '10000', '15000', '20000'],
        'Unnamed: 4': ['', '', '', '', ''],
        '3. Allocation %': ['', '', '22.2', '33.3', '44.5']
    }
    df1 = pd.DataFrame(sheet1_data)
    
    # Sheet 2: Financial Metrics with whitespace and unnamed columns
    sheet2_data = {
        '  Metric Name  ': ['Revenue', 'Profit', 'EPS', 'P/E Ratio'],
        'Unnamed: 1': ['', '', '', ''],
        'Q1 2024   ': ['1000000', '200000', '5.2', '18.5'],
        'Q2 2024': ['1200000', '250000', '6.1', '17.8'],
        'Unnamed: 4': ['', '', '', ''],
        '  Q3 2024': ['1500000', '300000', '7.3', '16.2']
    }
    df2 = pd.DataFrame(sheet2_data)
    
    # Sheet 3: Risk Data with only numeric prefixes
    sheet3_data = {
        '1. Risk Type': ['Market Risk', 'Credit Risk', 'Operational Risk'],
        '2. Probability': ['0.15', '0.08', '0.12'],
        '3. Impact Score': ['8', '6', '7'],
        '4. Mitigation Status': ['Active', 'Planned', 'Active']
    }
    df3 = pd.DataFrame(sheet3_data)
    
    # Sheet 4: Clean data (no unnamed columns)
    sheet4_data = {
        'Country': ['USA', 'UK', 'Germany', 'Japan'],
        'GDP (Trillion)': ['23.3', '3.1', '4.3', '4.9'],
        'Population (Million)': ['331', '67', '83', '125']
    }
    df4 = pd.DataFrame(sheet4_data)
    
    # Write to Excel
    test_file = 'test_messy_columns.xlsx'
    with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
        df1.to_excel(writer, sheet_name='Portfolio', index=False)
        df2.to_excel(writer, sheet_name='Financial Metrics', index=False)
        df3.to_excel(writer, sheet_name='Risk Analysis', index=False)
        df4.to_excel(writer, sheet_name='Country Data', index=False)
    
    print(f"✅ Created test file: {test_file}")
    return test_file


def test_column_cleaning(file_path):
    """Test column cleaning across all sheets."""
    
    print("\n" + "="*70)
    print("TESTING COLUMN CLEANING ACROSS ALL SHEETS")
    print("="*70)
    
    # Get all sheet names
    print("\n📋 Sheet Names:")
    sheets = get_sheet_names(file_path)
    for i, name in enumerate(sheets, 1):
        print(f"  {i}. {name}")
    
    # Read all sheets and check column cleaning
    print("\n" + "="*70)
    print("📚 READING AND CLEANING ALL SHEETS:")
    print("="*70)
    
    all_data = excel_reader_all_sheets(file_path)
    
    for sheet_name, df in all_data.items():
        print(f"\n📊 Sheet: '{sheet_name}'")
        print("-" * 70)
        
        if df is not None and not df.empty:
            print(f"   Shape: {df.shape[0]} rows × {df.shape[1]} columns")
            print(f"   Columns: {df.columns.tolist()}")
            
            # Check for issues
            issues = []
            for col in df.columns:
                if isinstance(col, str):
                    if col.strip().lower().startswith('unnamed'):
                        issues.append(f"❌ Unnamed column found: '{col}'")
                    if re.match(r'^\d+\.', col):
                        issues.append(f"❌ Numeric prefix found: '{col}'")
                    if col != col.strip():
                        issues.append(f"❌ Whitespace found: '{col}'")
            
            if issues:
                print(f"   ⚠️  ISSUES FOUND:")
                for issue in issues:
                    print(f"      {issue}")
            else:
                print(f"   ✅ All columns clean!")
            
            # Show first 3 rows
            print(f"\n   First 3 rows:")
            print(df.head(3).to_string(index=False))
        else:
            print(f"   ✗ Failed to read or empty")
    
    # Summary
    print("\n" + "="*70)
    print("✨ SUMMARY")
    print("="*70)
    successful = sum(1 for df in all_data.values() if df is not None and not df.empty)
    print(f"Successfully processed: {successful}/{len(sheets)} sheets")
    
    # Check if any unnamed columns remain
    has_unnamed = False
    for sheet_name, df in all_data.items():
        if df is not None:
            for col in df.columns:
                if isinstance(col, str) and 'unnamed' in col.lower():
                    has_unnamed = True
                    break
    
    if has_unnamed:
        print("❌ FAILED: Some unnamed columns still present")
    else:
        print("✅ SUCCESS: All unnamed columns removed!")


if __name__ == "__main__":
    import re
    
    # Create test file
    test_file = create_test_excel()
    
    # Test column cleaning
    test_column_cleaning(test_file)
    
    print("\n" + "="*70)
    print(f"📁 Test file saved as: {test_file}")
    print("="*70)
