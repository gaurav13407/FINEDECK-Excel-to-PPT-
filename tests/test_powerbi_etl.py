"""
Test Power BI ETL Pipeline
Tests the Excel to Power BI data model conversion
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.backend.app.services.powerbi_etl import ExcelToPowerBIProcessor, convert_excel_to_powerbi_model
import json


def test_powerbi_etl():
    """Test Power BI ETL with sample Excel file"""
    
    print("=" * 80)
    print("🧪 Testing Power BI ETL Pipeline")
    print("=" * 80)
    
    # Test with Sample PnL file
    excel_path = "examples/Sample_pnl.xlsx"
    
    if not Path(excel_path).exists():
        print(f"❌ Sample file not found: {excel_path}")
        return
    
    print(f"\n📂 Input: {excel_path}")
    print("-" * 80)
    
    # Run ETL pipeline
    try:
        powerbi_model = convert_excel_to_powerbi_model(excel_path)
        
        print("\n" + "=" * 80)
        print("📊 Power BI Model Summary")
        print("=" * 80)
        
        # Tables
        print(f"\n✅ Tables Created: {powerbi_model['metadata']['tables_count']}")
        for table_name, table_info in powerbi_model['model']['tables'].items():
            print(f"   • {table_name}")
            print(f"     - Type: {table_info['type']}")
            print(f"     - Rows: {table_info['row_count']}")
            print(f"     - Columns: {len(table_info['columns'])}")
            print(f"     - Columns: {', '.join(table_info['columns'][:5])}...")
            
            if table_info['hierarchies']:
                print(f"     - Hierarchies: {len(table_info['hierarchies'])}")
        
        # Relationships
        print(f"\n🔗 Relationships Detected: {powerbi_model['metadata']['relationships_count']}")
        for rel in powerbi_model['model']['relationships']:
            print(f"   • {rel['from_table']}[{rel['from_column']}] → {rel['to_table']}[{rel['to_column']}]")
        
        # DAX Measures
        print(f"\n📈 DAX Measures Generated: {powerbi_model['metadata']['measures_count']}")
        
        # Group measures by table
        measures_by_table = {}
        for measure in powerbi_model['model']['measures']:
            table = measure['table']
            if table not in measures_by_table:
                measures_by_table[table] = []
            measures_by_table[table].append(measure)
        
        for table_name, measures in measures_by_table.items():
            print(f"\n   Table: {table_name} ({len(measures)} measures)")
            for measure in measures[:3]:  # Show first 3
                print(f"     • {measure['name']} = {measure['expression']}")
            if len(measures) > 3:
                print(f"     ... and {len(measures) - 3} more")
        
        # Export sample
        print("\n" + "=" * 80)
        print("💾 Exporting Data Model")
        print("=" * 80)
        
        output_file = "examples/demo_PPT/powerbi_model.json"
        Path(output_file).parent.mkdir(parents=True, exist_ok=True)
        
        with open(output_file, 'w') as f:
            # Don't save actual data (too large)
            export_model = powerbi_model.copy()
            for table_name in export_model['model']['tables']:
                export_model['model']['tables'][table_name]['data'] = f"[{export_model['model']['tables'][table_name]['row_count']} rows]"
            
            json.dump(export_model, f, indent=2)
        
        print(f"✅ Model exported to: {output_file}")
        
        # Export CSV files
        processor = ExcelToPowerBIProcessor()
        processor.process_excel(excel_path)
        
        csv_output = "examples/demo_PPT/powerbi_csv"
        processor.export_to_csv(csv_output)
        
        print(f"✅ CSV files exported to: {csv_output}/")
        
        print("\n" + "=" * 80)
        print("✅ Test Complete!")
        print("=" * 80)
        
        print("\n📋 Next Steps:")
        print("   1. Check exported JSON model in: examples/demo_PPT/powerbi_model.json")
        print("   2. Check CSV files in: examples/demo_PPT/powerbi_csv/")
        print("   3. Import CSV files into Power BI Desktop")
        print("   4. Apply relationships and measures from JSON model")
        
    except Exception as e:
        print(f"\n❌ Error: {str(e)}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_powerbi_etl()
