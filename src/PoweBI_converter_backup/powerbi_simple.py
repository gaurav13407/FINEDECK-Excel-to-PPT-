"""
Ultra-Simple Power BI Export
Just creates CSV files + README instructions (no complex .pbit file)
"""

import json
from pathlib import Path
from typing import Dict, Any
import zipfile


def create_simple_package(dashboard_model: Dict[str, Any], csv_files: list, output_path: str) -> Path:
    """
    Create a simple ZIP package with:
    - CSV files (ready to import)
    - README with instructions
    - DAX measures text file
    
    User imports CSV manually - simple and works 100%
    """
    
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Create ZIP package
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        
        # 1. Add all CSV files
        for csv_file in csv_files:
            zipf.write(csv_file, arcname=f'data/{csv_file.name}')
        
        # 2. Add DAX measures
        measures_content = "Power BI DAX Measures\n" + "="*60 + "\n\n"
        measures_content += "Copy and paste these into Power BI Desktop:\n\n"
        
        for measure in dashboard_model['model'].get('measures', []):
            measures_content += f"{measure['name']} = {measure['expression']}\n\n"
        
        zipf.writestr('DAX_Measures.txt', measures_content)
        
        # 3. Add simple README
        readme = f"""# Power BI Dashboard Package

## Quick Start (3 Steps)

### Step 1: Open Power BI Desktop
- Download from: https://powerbi.microsoft.com/desktop/
- Launch the application

### Step 2: Import Data
1. Click "Get Data" → "Text/CSV"
2. Navigate to the `data/` folder
3. Select all CSV files
4. Click "Load"

### Step 3: Add Measures
1. Go to "Modeling" tab
2. Click "New Measure"
3. Open `DAX_Measures.txt`
4. Copy each measure and paste into Power BI

## What's Included

- **data/** - {len(csv_files)} CSV file(s) with clean data
- **DAX_Measures.txt** - {len(dashboard_model['model'].get('measures', []))} pre-built measures
- **README.txt** - This file

## Dashboard Info

- Template: {dashboard_model['metadata'].get('template_type', 'Custom')}
- Tables: {len(dashboard_model['model'].get('tables', {}))}
- Measures: {len(dashboard_model['model'].get('measures', []))}
- Relationships: {len(dashboard_model['model'].get('relationships', []))}

## Next Steps

After loading data and measures:
1. Create visuals (drag fields to canvas)
2. Add filters and slicers
3. Customize colors and formatting
4. Save as .pbix file
5. Publish to Power BI Service (optional)

## Support

For questions, contact support@findeck.com
"""
        
        zipf.writestr('README.txt', readme)
    
    print(f"✅ Created dashboard package: {output_file.name}")
    print(f"   📊 Size: {output_file.stat().st_size / 1024:.2f} KB")
    print(f"   📁 Contents: {len(csv_files)} CSV files + DAX measures + README")
    
    return output_file


def export_simple_package(excel_path: str, output_dir: str = None) -> Path:
    """
    One-function export: Excel → Simple ZIP package
    
    Usage:
        package = export_simple_package("sales.xlsx")
        # User downloads ZIP, extracts, imports CSV to Power BI
    """
    from src.PoweBI_converter.powerbi_etl import ExcelToPowerBIProcessor
    
    # Process Excel
    processor = ExcelToPowerBIProcessor()
    dashboard_model = processor.process_excel(excel_path)
    
    # Export CSV
    if output_dir is None:
        output_dir = Path(excel_path).parent / "powerbi_output"
    
    csv_dir = Path(output_dir) / "csv_data"
    processor.export_to_csv(str(csv_dir))
    
    # Get CSV files
    csv_files = list(csv_dir.glob("*.csv"))
    
    # Create package
    package_name = f"{Path(excel_path).stem}_powerbi_package.zip"
    package_path = Path(output_dir) / package_name
    
    return create_simple_package(dashboard_model, csv_files, str(package_path))


def create_pbit_file(dashboard_model: Dict[str, Any], csv_files: List[Path], output_path: str) -> Path:
    """
    Create a single .pbit (Power BI Template) file that users can open directly.
    
    SIMPLIFIED VERSION: Just creates the file structure, Power BI will handle the rest.
    """
    
    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Read CSV data
    import pandas as pd
    
    # Combine all CSV data
    embedded_data = {}
    for csv_file in csv_files:
        table_name = csv_file.stem
        df = pd.read_csv(csv_file)
        embedded_data[table_name] = {
            'columns': list(df.columns),
            'rows': df.to_dict('records')
        }
    
    # Simple Power BI template
    template = {
        'version': '2.0',
        'dataModel': {
            'name': 'Dashboard',
            'tables': dashboard_model['model'].get('tables', {}),
            'measures': dashboard_model['model'].get('measures', []),
            'relationships': dashboard_model['model'].get('relationships', [])
        },
        'layout': {
            'id': 1,
            'pages': [{
                'id': 1,
                'name': 'Report Page',
                'displayName': 'Dashboard',
                'width': 1280,
                'height': 720,
                'ordinal': 0
            }]
        },
        'metadata': dashboard_model.get('metadata', {})
    }
    
    # Create .pbit (just a ZIP file)
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as pbit:
        # Core files
        pbit.writestr('DataModelSchema', json.dumps(template['dataModel'], indent=2, cls=DateTimeEncoder))
        pbit.writestr('Report/Layout', json.dumps(template['layout'], indent=2))
        pbit.writestr('[Content_Types].xml', _get_content_types_xml())
        pbit.writestr('_rels/.rels', _get_rels_xml())
        
        # Add CSV data directly
        for csv_file in csv_files:
            pbit.write(csv_file, arcname=f'data/{csv_file.name}')
    
    print(f"✅ Created .pbit file: {output_file.name}")
    print(f"   📊 Size: {output_file.stat().st_size / 1024:.2f} KB")
    print(f"   📁 Tables: {len(template['dataModel'].get('tables', {}))} ")
    
    return output_file


def _create_default_visuals(dashboard_model: Dict[str, Any]) -> List[Dict]:
    """Create basic default visuals for the dashboard"""
    
    template_type = dashboard_model['metadata'].get('template_type', 'custom')
    tables = dashboard_model['model'].get('tables', [])
    measures = dashboard_model['model'].get('measures', [])
    
    visuals = []
    
    # Basic KPI cards for first 4 measures
    for i, measure in enumerate(measures[:4]):
        visuals.append({
            'type': 'card',
            'x': i * 250,
            'y': 20,
            'width': 200,
            'height': 100,
            'title': measure['name'].replace('Total', '').replace('Average', 'Avg'),
            'measure': measure['name']
        })
    
    # Add a table visual with data
    if tables and len(tables) > 0:
        first_table = tables[0] if isinstance(tables, list) else list(tables.values())[0]
        visuals.append({
            'type': 'table',
            'x': 20,
            'y': 150,
            'width': 600,
            'height': 400,
            'title': first_table['name'],
            'columns': [col['name'] for col in first_table.get('columns', [])[:6]]
        })
    
    # Add a chart if we have date column
    all_columns = []
    if isinstance(tables, list):
        for table in tables:
            all_columns.extend(table.get('columns', []))
    else:
        for table in tables.values():
            all_columns.extend(table.get('columns', []))
    
    date_columns = [col for col in all_columns if col.get('type') == 'datetime']
    if date_columns and measures:
        visuals.append({
            'type': 'lineChart',
            'x': 650,
            'y': 150,
            'width': 600,
            'height': 400,
            'title': 'Trend Over Time',
            'xAxis': date_columns[0]['name'],
            'yAxis': measures[0]['name']
        })
    
    return visuals


def _get_content_types_xml() -> str:
    """Get the required [Content_Types].xml for .pbit file"""
    return '''<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="json" ContentType="application/json"/>
  <Override PartName="/DataModelSchema" ContentType="application/json"/>
  <Override PartName="/Report/Layout" ContentType="application/json"/>
  <Override PartName="/version" ContentType="application/json"/>
</Types>'''


def _get_rels_xml() -> str:
    """Get the required _rels/.rels for .pbit file"""
    return '''<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="R1" Type="http://schemas.microsoft.com/sqlserver/dax/data" Target="/DataModelSchema"/>
  <Relationship Id="R2" Type="http://schemas.microsoft.com/sqlserver/reporting/reportlayout" Target="/Report/Layout"/>
</Relationships>'''


# Simple wrapper for easy use
def export_to_pbit(excel_path: str, output_dir: str = None) -> Path:
    """
    Simple one-function export: Excel → .pbit file
    
    Usage:
        pbit_file = export_to_pbit("sales_data.xlsx")
        # User can now open pbit_file directly in Power BI Desktop
    """
    from src.PoweBI_converter.powerbi_etl import ExcelToPowerBIProcessor
    
    # Process Excel
    processor = ExcelToPowerBIProcessor()
    dashboard_model = processor.process_excel(excel_path)
    
    # Export to CSV (temporary)
    if output_dir is None:
        output_dir = Path(excel_path).parent / "powerbi_output"
    
    csv_dir = Path(output_dir) / "csv_data"
    processor.export_to_csv(str(csv_dir))
    
    # Get CSV files
    csv_files = list(csv_dir.glob("*.csv"))
    
    # Create .pbit file
    pbit_filename = f"{Path(excel_path).stem}_dashboard.pbit"
    pbit_path = Path(output_dir) / pbit_filename
    
    return create_pbit_file(dashboard_model, csv_files, str(pbit_path))
