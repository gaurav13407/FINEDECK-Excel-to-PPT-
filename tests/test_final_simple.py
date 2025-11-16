"""
Test simple ZIP package (CSV + README)
"""

import sys
from pathlib import Path

project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src" / "backend"))

import pandas as pd
from app.services.powerbi_export import export_simple_package


print("\n" + "="*60)
print("SIMPLE POWER BI PACKAGE TEST")
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

# Generate package
print("\n[2] Generating ZIP package...")
package = export_simple_package(str(excel_file), str(output_dir))

print("\n[3] Package created!")
print(f"    File: {package.name}")
print(f"    Size: {package.stat().st_size / 1024:.2f} KB")
print(f"    Path: {package}")

print("\n" + "="*60)
print("SUCCESS!")
print("="*60)
print("User can now:")
print("  1. Download ZIP file")
print("  2. Extract it")
print("  3. Open Power BI Desktop")
print("  4. Import CSV files")
print("  5. Copy DAX measures")
print("  6. Done!")
print("="*60)
