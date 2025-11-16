"""
Test Power BI Dashboard Package Download System

This script tests the complete workflow:
1. Create sample Excel file
2. Process with ETL pipeline
3. Generate complete package (ZIP)
4. Verify package contents
"""

import sys
from pathlib import Path

# Add project root to Python path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root / "src" / "backend"))

import pandas as pd
import json
import zipfile
from datetime import datetime, timedelta

from app.services.powerbi_etl import ExcelToPowerBIProcessor
from app.services.powerbi_template_generator import create_dashboard_package


def create_test_sales_excel():
    """Create test Excel file with sales data"""
    
    # Generate 50 sales transactions
    dates = pd.date_range(start='2024-01-01', end='2024-03-31', freq='D')
    regions = ['North', 'South', 'East', 'West', 'Central']
    products = ['Laptop', 'Phone', 'Tablet', 'Monitor', 'Keyboard']
    
    data = []
    for i in range(50):
        data.append({
            'Date': dates[i % len(dates)],
            'Region': regions[i % len(regions)],
            'Product': products[i % len(products)],
            'Sales Amount': (i + 1) * 5000 + (i % 10) * 1000,
            'Quantity Sold': (i % 20) + 1,
            'Sales Target': (i + 1) * 6000,
            'Salesperson': f"Rep_{(i % 10) + 1}",
            'Customer Segment': ['Enterprise', 'SMB', 'Consumer'][i % 3]
        })
    
    df = pd.DataFrame(data)
    
    output_dir = project_root / "examples" / "test_data" / "package_test"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    excel_path = output_dir / "sales_test.xlsx"
    df.to_excel(excel_path, index=False, sheet_name='Sales Data')
    
    print(f"✅ Created test Excel: {excel_path}")
    print(f"   📊 Rows: {len(df)}, Columns: {len(df.columns)}")
    
    return excel_path


def test_package_creation():
    """Test complete package creation workflow"""
    
    print("\n" + "="*70)
    print("🧪 TESTING POWER BI PACKAGE DOWNLOAD SYSTEM")
    print("="*70)
    
    # Step 1: Create test Excel file
    print("\n📁 Step 1: Creating test Excel file...")
    excel_path = create_test_sales_excel()
    
    # Step 2: Process with ETL pipeline
    print("\n⚙️ Step 2: Processing with ETL pipeline...")
    processor = ExcelToPowerBIProcessor()
    processor.template_type = 'sales_performance'  # Set template type
    powerbi_model = processor.process_excel(str(excel_path))
    
    print(f"   ✅ Tables: {powerbi_model['metadata']['tables_count']}")
    print(f"   ✅ Measures: {powerbi_model['metadata']['measures_count']}")
    print(f"   ✅ Relationships: {powerbi_model['metadata']['relationships_count']}")
    
    # Step 3: Export CSV files
    print("\n📊 Step 3: Exporting CSV files...")
    output_dir = excel_path.parent / "powerbi_output"
    csv_dir = output_dir / "csv_data"
    processor.export_to_csv(str(csv_dir))
    
    csv_files = list(csv_dir.glob("*.csv"))
    print(f"   ✅ Exported {len(csv_files)} CSV files:")
    for csv_file in csv_files:
        df = pd.read_csv(csv_file)
        print(f"      • {csv_file.name}: {len(df)} rows, {len(df.columns)} columns")
    
    # Step 4: Create dashboard package
    print("\n📦 Step 4: Creating dashboard package...")
    package_zip = create_dashboard_package(
        dashboard_model=powerbi_model,
        dashboard_type='sales_performance',
        csv_files=csv_files,
        output_dir=output_dir
    )
    
    print(f"   ✅ Package created: {package_zip.name}")
    print(f"   ✅ Package size: {package_zip.stat().st_size / 1024:.2f} KB")
    
    # Step 5: Verify package contents
    print("\n🔍 Step 5: Verifying package contents...")
    
    with zipfile.ZipFile(package_zip, 'r') as zip_ref:
        file_list = zip_ref.namelist()
        print(f"   ✅ Total files in ZIP: {len(file_list)}")
        print("\n   📂 Package Structure:")
        
        for file_name in sorted(file_list):
            file_info = zip_ref.getinfo(file_name)
            size_kb = file_info.file_size / 1024
            print(f"      • {file_name} ({size_kb:.2f} KB)")
        
        # Check required components
        required_files = {
            'CSV data': any('data/' in f and f.endswith('.csv') for f in file_list),
            'Dashboard model': 'dashboard_model.json' in file_list,
            'DAX measures': 'DAX_Measures.txt' in file_list,
            'Quick Start guide': 'QUICK_START.md' in file_list
        }
        
        print("\n   ✅ Required Components Check:")
        for component, exists in required_files.items():
            status = "✅" if exists else "❌"
            print(f"      {status} {component}")
        
        # Verify JSON model
        if 'dashboard_model.json' in file_list:
            model_json = zip_ref.read('dashboard_model.json')
            model_data = json.loads(model_json)
            print("\n   📊 Dashboard Model Details:")
            print(f"      • Tables: {len(model_data['model']['tables'])}")
            print(f"      • Relationships: {len(model_data['model']['relationships'])}")
            print(f"      • Measures: {len(model_data['model']['measures'])}")
        
        # Verify DAX measures
        if 'DAX_Measures.txt' in file_list:
            dax_content = zip_ref.read('DAX_Measures.txt').decode('utf-8')
            measure_count = dax_content.count('MEASURE')
            print(f"\n   📈 DAX Measures File:")
            print(f"      • Total measures: {measure_count}")
            print(f"      • File size: {len(dax_content)} characters")
        
        # Verify Quick Start guide
        if 'QUICK_START.md' in file_list:
            guide_content = zip_ref.read('QUICK_START.md').decode('utf-8')
            print(f"\n   📖 Quick Start Guide:")
            print(f"      • File size: {len(guide_content)} characters")
            print(f"      • Contains steps: {'Step 1' in guide_content}")
    
    # Step 6: Final verdict
    print("\n" + "="*70)
    print("🎯 PACKAGE DOWNLOAD TEST RESULTS")
    print("="*70)
    
    all_checks_passed = all(required_files.values())
    
    if all_checks_passed:
        print("\n✅ SUCCESS: Complete package system working perfectly!")
        print("\n📦 Package Contents:")
        print("   ✅ CSV data files (ready for Power BI import)")
        print("   ✅ Dashboard model JSON (metadata)")
        print("   ✅ DAX measures file (copy-paste ready)")
        print("   ✅ Quick Start guide (step-by-step instructions)")
        print("\n🚀 Users can now:")
        print("   1. Download the ZIP package")
        print("   2. Extract files")
        print("   3. Follow Quick Start guide")
        print("   4. Import CSV into Power BI Desktop")
        print("   5. Dashboard ready in 5 minutes!")
    else:
        print("\n❌ FAILED: Some required components missing")
        print("\n   Missing:")
        for component, exists in required_files.items():
            if not exists:
                print(f"   ❌ {component}")
    
    print("\n" + "="*70)
    
    return package_zip if all_checks_passed else None


if __name__ == "__main__":
    try:
        package = test_package_creation()
        
        if package:
            print(f"\n✅ Package ready for download: {package}")
            print(f"   📍 Location: {package.parent}")
            print(f"   💾 Size: {package.stat().st_size / 1024:.2f} KB")
    
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()
