"""
Comprehensive Power BI Dashboard Review & Testing
Tests with multiple sample files to verify ETL quality
"""

import sys
from pathlib import Path
import pandas as pd

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.PoweBI_converter.powerbi_etl import ExcelToPowerBIProcessor
import json


def create_test_excel_files():
    """Create clean test Excel files for different dashboard types"""
    
    output_dir = Path("examples/test_data")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    print("📝 Creating test Excel files...")
    print("=" * 80)
    
    # 1. Sales Performance Data
    sales_data = pd.DataFrame({
        'Date': pd.date_range('2024-01-01', periods=12, freq='MS'),
        'Region': ['North', 'South', 'East', 'West'] * 3,
        'Product': ['Laptop', 'Phone', 'Tablet', 'Desktop'] * 3,
        'Sales': [50000, 35000, 28000, 42000, 55000, 38000, 30000, 45000, 
                 60000, 40000, 32000, 48000],
        'Quantity': [50, 120, 80, 30, 55, 125, 85, 32, 60, 130, 90, 35],
        'Target': [48000, 36000, 27000, 40000, 52000, 39000, 29000, 43000,
                  58000, 41000, 31000, 46000]
    })
    sales_file = output_dir / "sales_performance.xlsx"
    sales_data.to_excel(sales_file, sheet_name='Sales', index=False)
    print(f"✅ Created: {sales_file}")
    
    # 2. Financial KPI Data
    financial_data = pd.DataFrame({
        'Month': pd.date_range('2024-01-01', periods=12, freq='MS'),
        'Revenue': [100000, 120000, 115000, 130000, 140000, 135000,
                   150000, 145000, 160000, 170000, 165000, 180000],
        'Expenses': [70000, 80000, 76000, 85000, 90000, 88000,
                    95000, 92000, 100000, 105000, 102000, 110000],
        'Profit': [30000, 40000, 39000, 45000, 50000, 47000,
                  55000, 53000, 60000, 65000, 63000, 70000],
        'Budget': [95000, 115000, 110000, 125000, 135000, 130000,
                  145000, 140000, 155000, 165000, 160000, 175000]
    })
    financial_file = output_dir / "financial_kpi.xlsx"
    financial_data.to_excel(financial_file, sheet_name='Financials', index=False)
    print(f"✅ Created: {financial_file}")
    
    # 3. Marketing Analytics Data
    marketing_data = pd.DataFrame({
        'Date': pd.date_range('2024-01-01', periods=10, freq='W'),
        'Campaign': ['Google Ads', 'Facebook', 'LinkedIn', 'Instagram', 'Twitter'] * 2,
        'Impressions': [100000, 80000, 60000, 90000, 50000, 110000, 85000, 65000, 95000, 55000],
        'Clicks': [5000, 4000, 3000, 4500, 2500, 5500, 4200, 3200, 4800, 2800],
        'Conversions': [250, 200, 150, 225, 125, 275, 210, 160, 240, 140],
        'Spend': [10000, 8000, 6000, 9000, 5000, 11000, 8500, 6500, 9500, 5500]
    })
    marketing_file = output_dir / "marketing_analytics.xlsx"
    marketing_data.to_excel(marketing_file, sheet_name='Campaigns', index=False)
    print(f"✅ Created: {marketing_file}")
    
    # 4. Revenue & Profit by Category
    revenue_data = pd.DataFrame({
        'Date': pd.date_range('2024-01-01', periods=8, freq='Q'),
        'Category': ['Electronics', 'Clothing', 'Food', 'Furniture'] * 2,
        'Revenue': [500000, 300000, 200000, 400000, 550000, 320000, 220000, 420000],
        'Cost': [350000, 180000, 120000, 280000, 385000, 192000, 132000, 294000],
        'Units_Sold': [5000, 12000, 8000, 2000, 5500, 13000, 8500, 2100]
    })
    revenue_file = output_dir / "revenue_profit.xlsx"
    revenue_data.to_excel(revenue_file, sheet_name='Revenue', index=False)
    print(f"✅ Created: {revenue_file}")
    
    # 5. Operations Efficiency
    operations_data = pd.DataFrame({
        'Date': pd.date_range('2024-01-01', periods=12, freq='MS'),
        'Production': [10000, 10500, 11000, 10800, 11500, 11200,
                      12000, 11800, 12500, 12200, 13000, 12800],
        'Capacity': [12000, 12000, 12000, 12000, 13000, 13000,
                    13000, 13000, 14000, 14000, 14000, 14000],
        'Downtime_Hours': [24, 18, 20, 22, 16, 19, 15, 17, 14, 16, 12, 15],
        'Resource_Count': [50, 50, 52, 52, 55, 55, 58, 58, 60, 60, 62, 62]
    })
    operations_file = output_dir / "operations_efficiency.xlsx"
    operations_data.to_excel(operations_file, sheet_name='Operations', index=False)
    print(f"✅ Created: {operations_file}")
    
    print("\n" + "=" * 80)
    return [sales_file, financial_file, marketing_file, revenue_file, operations_file]


def test_powerbi_dashboard(excel_path: Path, expected_template: str):
    """Test Power BI ETL with a specific file"""
    
    print(f"\n{'=' * 80}")
    print(f"📊 Testing: {excel_path.name}")
    print(f"Expected Template: {expected_template}")
    print("=" * 80)
    
    try:
        # Run ETL
        processor = ExcelToPowerBIProcessor()
        model = processor.process_excel(str(excel_path))
        
        # Extract summary
        tables = model['model']['tables']
        relationships = model['model']['relationships']
        measures = model['model']['measures']
        
        print(f"\n✅ ETL COMPLETE")
        print(f"   📋 Tables: {len(tables)}")
        print(f"   🔗 Relationships: {len(relationships)}")
        print(f"   📊 Measures: {len(measures)}")
        
        # Detailed table analysis
        print(f"\n📋 TABLE DETAILS:")
        for table_name, table_info in tables.items():
            print(f"\n   • {table_name}")
            print(f"     Type: {table_info['type']}")
            print(f"     Rows: {table_info['row_count']}")
            print(f"     Columns ({len(table_info['columns'])}): {', '.join(table_info['columns'][:8])}")
            
            # Show sample data
            sample_data = table_info['data'][:3] if table_info['data'] != f"[{table_info['row_count']} rows]" else []
            if sample_data:
                print(f"     Sample: {sample_data[0]}")
        
        # Relationship quality check
        if relationships:
            print(f"\n🔗 RELATIONSHIPS:")
            for rel in relationships[:5]:  # Show first 5
                print(f"   • {rel['from_table']}[{rel['from_column']}] → {rel['to_table']}[{rel['to_column']}]")
            if len(relationships) > 5:
                print(f"   ... and {len(relationships) - 5} more")
        
        # Measure quality check
        if measures:
            print(f"\n📊 DAX MEASURES:")
            measures_by_type = {}
            for measure in measures:
                measure_type = measure['name'].split(measure['table'])[-1]  # Extract measure type
                if measure_type not in measures_by_type:
                    measures_by_type[measure_type] = []
                measures_by_type[measure_type].append(measure)
            
            for measure_type, measure_list in list(measures_by_type.items())[:5]:
                print(f"   • {measure_type}: {len(measure_list)} measures")
                print(f"     Example: {measure_list[0]['expression']}")
        
        # Data quality assessment
        print(f"\n🔍 DATA QUALITY:")
        
        # Check for unnamed columns
        unnamed_count = sum(1 for table in tables.values() 
                          for col in table['columns'] if 'unnamed' in col.lower())
        if unnamed_count > 0:
            print(f"   ⚠️  Warning: {unnamed_count} unnamed columns detected")
        else:
            print(f"   ✅ All columns properly named")
        
        # Check for numeric columns
        total_numeric = sum(len([col for col in table['columns'] 
                                if any(num in col.lower() for num in ['revenue', 'sales', 'cost', 'profit', 'quantity', 'impressions', 'clicks', 'conversions', 'spend', 'production', 'capacity'])]) 
                          for table in tables.values())
        print(f"   ✅ {total_numeric} numeric/metric columns identified")
        
        # Check for date columns
        has_date = any(len(table['hierarchies']) > 0 for table in tables.values())
        if has_date:
            print(f"   ✅ Date hierarchies created")
        else:
            print(f"   ⚠️  No date columns detected")
        
        # Template detection test
        from src.backend.app.api.v1.endpoints.powerbi import _detect_template
        detected_template = _detect_template(model)
        
        print(f"\n🎯 TEMPLATE DETECTION:")
        print(f"   Expected: {expected_template}")
        print(f"   Detected: {detected_template}")
        if detected_template == expected_template:
            print(f"   ✅ MATCH!")
        else:
            print(f"   ⚠️  MISMATCH - Review detection logic")
        
        # Export for review
        output_dir = excel_path.parent / "powerbi_output"
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Export JSON model
        json_file = output_dir / f"{excel_path.stem}_model.json"
        with open(json_file, 'w') as f:
            # Don't save full data (too large)
            export_model = model.copy()
            for table_name in export_model['model']['tables']:
                export_model['model']['tables'][table_name]['data'] = f"[{export_model['model']['tables'][table_name]['row_count']} rows]"
            json.dump(export_model, f, indent=2)
        
        # Export CSV
        csv_dir = output_dir / excel_path.stem
        processor.export_to_csv(str(csv_dir))
        
        print(f"\n💾 EXPORTED:")
        print(f"   JSON: {json_file}")
        print(f"   CSV:  {csv_dir}/")
        
        return {
            'success': True,
            'tables_count': len(tables),
            'relationships_count': len(relationships),
            'measures_count': len(measures),
            'template_match': detected_template == expected_template
        }
        
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        import traceback
        traceback.print_exc()
        return {'success': False, 'error': str(e)}


def main():
    """Run comprehensive Power BI dashboard tests"""
    
    print("=" * 80)
    print("🧪 COMPREHENSIVE POWER BI DASHBOARD REVIEW")
    print("=" * 80)
    
    # Create test files
    test_files = create_test_excel_files()
    
    # Define expected templates for each file
    test_cases = [
        (test_files[0], 'sales_performance'),
        (test_files[1], 'financial_kpi'),
        (test_files[2], 'marketing_analytics'),
        (test_files[3], 'revenue_profit'),
        (test_files[4], 'operations_efficiency')
    ]
    
    # Run tests
    results = []
    for excel_file, expected_template in test_cases:
        result = test_powerbi_dashboard(excel_file, expected_template)
        results.append({
            'file': excel_file.name,
            'template': expected_template,
            **result
        })
    
    # Summary report
    print("\n" + "=" * 80)
    print("📊 TEST SUMMARY")
    print("=" * 80)
    
    for result in results:
        status = "✅ PASS" if result['success'] else "❌ FAIL"
        template_status = "✅" if result.get('template_match', False) else "⚠️"
        
        print(f"\n{status} {result['file']}")
        if result['success']:
            print(f"   Tables: {result['tables_count']}, Relationships: {result['relationships_count']}, Measures: {result['measures_count']}")
            print(f"   {template_status} Template: {result['template']}")
        else:
            print(f"   Error: {result.get('error', 'Unknown')}")
    
    # Final verdict
    success_count = sum(1 for r in results if r['success'])
    print(f"\n{'=' * 80}")
    print(f"✅ {success_count}/{len(results)} tests passed")
    print("=" * 80)
    
    print("\n📋 NEXT STEPS:")
    print("   1. Review exported JSON models in examples/test_data/powerbi_output/")
    print("   2. Check CSV files for data quality")
    print("   3. Import CSV into Power BI Desktop to verify")
    print("   4. Test API endpoint with Postman/curl")


if __name__ == "__main__":
    main()
