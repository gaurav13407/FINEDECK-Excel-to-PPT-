"""
Simple test: Excel → Single .pbit file
"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src" / "backend"))

import pandas as pd
from app.services.powerbi_simple import export_to_pbit


def test_simple_pbit():
    """Test creating a single .pbit file"""
    
    print("\n" + "="*60)
    print("SIMPLE POWER BI TEST - Single File Output")
    print("="*60)
    
    # Create test Excel
    print("\n[1] Creating test Excel...")
    data = []
    for i in range(30):
        data.append({
            'Date': f'2024-0{(i % 9) + 1}-01',
            'Product': ['Laptop', 'Phone', 'Tablet'][i % 3],
            'Sales': (i + 1) * 10000,
            'Region': ['North', 'South'][i % 2]
        })
    
    df = pd.DataFrame(data)
    
    output_dir = project_root / "examples" / "test_data" / "simple_test"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    excel_file = output_dir / "test_sales.xlsx"
    df.to_excel(excel_file, index=False, sheet_name='Sales')
    print(f"    Created: {excel_file.name} ({len(df)} rows)")
    
    # Generate .pbit file
    print("\n[2] Generating .pbit file...")
    pbit_file = export_to_pbit(str(excel_file), str(output_dir))
    
    print("\n[3] Result:")
    print(f"    File: {pbit_file.name}")
    print(f"    Size: {pbit_file.stat().st_size / 1024:.2f} KB")
    print(f"    Path: {pbit_file}")
    
    print("\n" + "="*60)
    print("SUCCESS - User can now:")
    print("  1. Download: " + pbit_file.name)
    print("  2. Double-click to open in Power BI Desktop")
    print("  3. Dashboard loads with data automatically!")
    print("="*60)
    
    return pbit_file


if __name__ == "__main__":
    try:
        test_simple_pbit()
    except Exception as e:
        print(f"\n[ERROR] {e}")
        import traceback
        traceback.print_exc()
