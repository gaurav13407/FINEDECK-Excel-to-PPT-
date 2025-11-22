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
