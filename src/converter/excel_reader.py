"""
Excel Reader Module for FinDeck Converter
==========================================

This module provides robust Excel file reading capabilities with:
- Automatic header detection
- Column cleaning and sanitization  
- Support for multiple sheets
- Handling of unnamed/malformed columns
- Numeric column type inference
"""

import sys
import re
import pandas as pd
import openpyxl


def excel_reader(path: str, sheet: str | int | None = None) -> pd.DataFrame:
    """
    Read an Excel file and return a cleaned DataFrame.
    
    Args:
        path: Path to the Excel file
        sheet: Sheet name (str), sheet index (int), or None for first sheet
        
    Returns:
        Cleaned pandas DataFrame
        
    Raises:
        Exception: If file cannot be read or processed
    """
    try:
        # Determine sheet to read
        if sheet is None:
            sheet_page = 0
        else:
            sheet_page = sheet

        # Read raw Excel data
        raw_df = pd.read_excel(path, sheet_name=sheet_page)
        
        if raw_df.empty:
            return pd.DataFrame()

        # Check if there are any 'Unnamed' columns
        has_unnamed = any(
            isinstance(c, str) and c.strip().lower().startswith("unnamed")
            for c in raw_df.columns
        )

        if not has_unnamed:
            # No Unnamed headers => skip header promotion
            # Just clean and return
            df = clean_columns(raw_df)
            df = df.dropna(how="all").reset_index(drop=True)
            return df
        
        # Find the actual header row
        header_row = find_header_row(raw_df, required_keywords=None)
        
        # If header is at or beyond the last row, return empty DataFrame with that row as columns
        if header_row >= raw_df.shape[0] - 1:
            header_value = raw_df.iloc[header_row].fillna("").tolist()
            df = pd.DataFrame(columns=header_value)
            return df
        
        # Promote the header row
        df = raw_df.iloc[header_row + 1:].reset_index(drop=True)
        df.columns = raw_df.iloc[header_row].fillna("").tolist()
        
        # Clean columns
        df = clean_columns(df)
        
        # Drop empty rows
        df = df.dropna(how="all").reset_index(drop=True)
        
        # Try to infer numeric columns
        for col in df.columns:
            try:
                sample = df[col].astype(str).str.strip().replace({"nan": ""})
                non_empty = sample[sample != ""]
                
                if len(non_empty) == 0:
                    continue
                
                # Check if at least 40% of non-empty values look numeric
                num_like = non_empty.str.match(r'^-?\d+(\.\d+)?$').sum()
                if num_like >= max(1, int(0.4 * len(non_empty))):
                    df[col] = pd.to_numeric(df[col], errors="coerce")
            except Exception:
                continue
        
        return df

    except Exception as e:
        raise Exception(f"Error reading Excel file: {str(e)}")


def find_header_row(df: pd.DataFrame, required_keywords=None) -> int:
    """
    Find the first row that contains at least one of the required keywords.
    
    Args:
        df: DataFrame to search
        required_keywords: List of keywords to look for (default: common financial terms)
        
    Returns:
        Index of the header row (0 if not found)
    """
    try:
        if required_keywords is None:
            required_keywords = [
                "asset", "value", "sector", "country", "ticker", "symbol",
                "amount", "price", "name", "date", "category", "type",
                "description", "quantity", "total", "balance", "account"
            ]
        
        required_keywords = [k.lower() for k in required_keywords]
        
        for idx, row in df.iterrows():
            # Convert row to lowercase strings, ignoring NaN
            tokens = row.astype(str).str.lower().fillna("")
            
            # Count matches
            matches = sum(
                any(k in cell for k in required_keywords) 
                for cell in tokens
            )
            
            # If at least 1 match found, consider this the header row
            if matches >= 1:
                return idx
        
        # No header found, assume row 0
        return 0
        
    except Exception:
        return 0


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean DataFrame columns by:
    - Dropping fully empty columns
    - Removing 'Unnamed' columns
    - Stripping whitespace
    - Removing numeric prefixes (e.g., '2. Column' -> 'Column')
    
    Args:
        df: DataFrame to clean
        
    Returns:
        Cleaned DataFrame
    """
    # Drop fully empty columns
    df = df.dropna(axis=1, how="all")
    
    # Check if there are unnamed columns
    has_unnamed = any(
        isinstance(c, str) and c.strip().lower().startswith("unnamed")
        for c in df.columns
    )
    
    if not has_unnamed:
        # Just strip whitespace from column names
        new_cols = []
        for c in df.columns:
            if isinstance(c, str):
                c = c.strip()
                # Remove numeric prefixes like '2. Column Name'
                c = re.sub(r'^\d+\.\s*', '', c)
            new_cols.append(c)
        df.columns = new_cols
        return df
    
    # Drop columns whose names start with 'Unnamed'
    cols_to_keep = [
        not (isinstance(c, str) and c.strip().lower().startswith("unnamed"))
        for c in df.columns
    ]
    df = df.loc[:, cols_to_keep]
    
    # Clean remaining column names
    new_cols = []
    for c in df.columns:
        if isinstance(c, str):
            c = c.strip()
            # Remove numeric prefixes like '2. Portfolio Allocation Data'
            c = re.sub(r'^\d+\.\s*', '', c)
        new_cols.append(c)
    df.columns = new_cols
    
    return df


def get_sheet_names(path: str)->list:
    """
    Get all sheet names from an excel file."""
    try:
        xl_file=pd.ExcelFile(path)
        return xl_file.sheet_names
    except Exception as e:
        raise Exception(f"Error getting sheet names: {str(e)}")



def excel_reader_all_sheets(path: str)->dict:
    """
    Read all sheets from an excel file and return a dictionary of DataFrames."""
    try:
        sheet_name=get_sheet_names(path)
        result={}
        for sheet_name in sheet_name:
            try:
                df=excel_reader(path,sheet=sheet_name)
                result[sheet_name]=df
            except Exception as e:
                print(f"Warning reading sheet {sheet_name}: {str(e)}")
                result[sheet_name]=None
        return result
    except Exception as e:
        raise Exception(f"Error reading all sheets: {str(e)}")


# Test/example usage
if __name__ == "__main__":
    import os
    
    # Test with a sample file if it exists
    test_file = "example\Company_Data\AAPL_Financial_Data.xlsx"
    if os.path.exists(test_file):
        print(f"Testing with: {test_file}")
        print("\n" + "="*60)
        
        # Test 1: Get all sheet names
        print("\n📋 Sheet Names:")
        sheets = get_sheet_names(test_file)
        for i, name in enumerate(sheets, 1):
            print(f"  {i}. {name}")
        
        # Test 2: Read first sheet only
        print("\n" + "="*60)
        print("\n📊 Reading First Sheet:")
        df = excel_reader(test_file, sheet=None)
        print(f"Shape: {df.shape}")
        print(f"Columns: {df.columns.tolist()}")
        print(f"\nFirst 3 rows:\n{df.head(3)}")
        
        # Test 3: Read specific sheet by name
        if len(sheets) > 1:
            print("\n" + "="*60)
            print(f"\n📊 Reading Sheet '{sheets[1]}':")
            df2 = excel_reader(test_file, sheet=sheets[1])
            print(f"Shape: {df2.shape}")
            print(f"Columns: {df2.columns.tolist()}")
        
        # Test 4: Read ALL sheets
        print("\n" + "="*60)
        print("\n📚 Reading ALL Sheets:")
        all_data = excel_reader_all_sheets(test_file)
        for sheet_name, df in all_data.items():
            if df is not None:
                print(f"  ✓ {sheet_name}: {df.shape[0]} rows × {df.shape[1]} columns")
            else:
                print(f"  ✗ {sheet_name}: Failed to read")
        
    else:
        print(f"Test file not found: {test_file}")
        print("\nYou can test with any Excel file by running:")
        print("  python excel_reader.py")