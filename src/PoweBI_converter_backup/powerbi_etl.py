"""
Power BI ETL Pipeline - Excel to Power BI Data Model Converter
Automatically cleans, normalizes, and structures Excel data for Power BI.
Removes the need for Power BI Desktop entirely.
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Any, Tuple, Optional
from pathlib import Path
import re
from datetime import datetime
from collections import defaultdict


class ExcelToPowerBIProcessor:
    """
    Automated ETL pipeline that converts raw Excel into Power BI-ready data models.
    
    Features:
    - Data cleaning (nulls, duplicates, type inference)
    - Column normalization (standardized naming)
    - Relationship detection (foreign keys, hierarchies)
    - Fact/dimension table creation
    - DAX measure generation
    """
    
    def __init__(self):
        self.raw_data: Dict[str, pd.DataFrame] = {}
        self.cleaned_data: Dict[str, pd.DataFrame] = {}
        self.data_model: Dict[str, Any] = {}
        self.relationships: List[Dict[str, str]] = []
        self.dax_measures: List[Dict[str, str]] = []
        
    def process_excel(self, excel_path: str) -> Dict[str, Any]:
        """
        Main pipeline: Excel → Clean → Model → Relationships → DAX
        
        Args:
            excel_path: Path to Excel file
            
        Returns:
            Complete Power BI data model specification
        """
        print(f"🔄 Starting Power BI ETL Pipeline...")
        
        # Step 1: Load Excel
        self._load_excel(excel_path)
        
        # Step 2: Clean data
        self._clean_data()
        
        # Step 3: Normalize columns
        self._normalize_columns()
        
        # Step 4: Build data model (fact/dimension tables)
        self._build_data_model()
        
        # Step 5: Detect relationships
        self._auto_create_relationships()
        
        # Step 6: Generate DAX measures
        self._generate_dax_measures()
        
        # Step 7: Package for Power BI
        result = self._package_model()
        
        print(f"✅ Power BI model ready with {len(self.data_model['tables'])} tables, "
              f"{len(self.relationships)} relationships, {len(self.dax_measures)} measures")
        
        return result
    
    def _load_excel(self, excel_path: str):
        """Load all sheets from Excel file"""
        print(f"📂 Loading Excel: {excel_path}")
        
        excel_file = pd.ExcelFile(excel_path)
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            if not df.empty:
                self.raw_data[sheet_name] = df
                print(f"   ✓ Loaded '{sheet_name}': {df.shape[0]} rows, {df.shape[1]} columns")
    
    def _clean_data(self):
        """Clean raw data: handle nulls, duplicates, data types"""
        print(f"\n🧹 Cleaning data...")
        
        for sheet_name, df in self.raw_data.items():
            df_clean = df.copy()
            
            # Remove completely empty rows/columns
            df_clean = df_clean.dropna(how='all', axis=0)
            df_clean = df_clean.dropna(how='all', axis=1)
            
            # Remove duplicate rows
            initial_rows = len(df_clean)
            df_clean = df_clean.drop_duplicates()
            removed_dupes = initial_rows - len(df_clean)
            
            # Infer and fix data types
            for col in df_clean.columns:
                # Try to convert to numeric
                if df_clean[col].dtype == 'object':
                    try:
                        df_clean[col] = pd.to_numeric(df_clean[col], errors='ignore')
                    except:
                        pass
                
                # Try to convert to datetime
                if df_clean[col].dtype == 'object':
                    try:
                        df_clean[col] = pd.to_datetime(df_clean[col], errors='ignore')
                    except:
                        pass
            
            self.cleaned_data[sheet_name] = df_clean
            print(f"   ✓ '{sheet_name}': Removed {removed_dupes} duplicates, "
                  f"Cleaned {df_clean.shape[0]} rows")
    
    def _normalize_columns(self):
        """Standardize column names for Power BI compatibility"""
        print(f"\n🔧 Normalizing column names...")
        
        for sheet_name, df in self.cleaned_data.items():
            new_columns = []
            
            for col in df.columns:
                # Convert to string
                col_str = str(col)
                
                # Remove special characters, keep alphanumeric and spaces
                col_normalized = re.sub(r'[^a-zA-Z0-9\s]', '', col_str)
                
                # Replace multiple spaces with single space
                col_normalized = re.sub(r'\s+', ' ', col_normalized)
                
                # Trim and title case
                col_normalized = col_normalized.strip().title()
                
                # Remove spaces for Power BI column names
                col_normalized = col_normalized.replace(' ', '')
                
                # Ensure unique column names
                if col_normalized in new_columns:
                    counter = 1
                    while f"{col_normalized}{counter}" in new_columns:
                        counter += 1
                    col_normalized = f"{col_normalized}{counter}"
                
                new_columns.append(col_normalized)
            
            df.columns = new_columns
            self.cleaned_data[sheet_name] = df
            print(f"   ✓ '{sheet_name}': Normalized {len(new_columns)} columns")
    
    def _build_data_model(self):
        """Create fact and dimension tables for Star Schema"""
        print(f"\n🏗️  Building data model...")
        
        tables = {}
        
        for sheet_name, df in self.cleaned_data.items():
            # Classify table type (Fact or Dimension)
            table_type = self._classify_table_type(df, sheet_name)
            
            # Add index column if not present
            if 'Id' not in df.columns:
                df.insert(0, f'{sheet_name}Id', range(1, len(df) + 1))
            
            tables[sheet_name] = {
                'name': sheet_name,
                'type': table_type,
                'data': df,
                'columns': list(df.columns),
                'row_count': len(df),
                'measures': [],
                'hierarchies': []
            }
            
            # Detect hierarchies (e.g., Year -> Quarter -> Month)
            hierarchies = self._detect_hierarchies(df)
            if hierarchies:
                tables[sheet_name]['hierarchies'] = hierarchies
            
            print(f"   ✓ '{sheet_name}': {table_type} table with {len(df)} rows")
        
        self.data_model = {'tables': tables}
    
    def _classify_table_type(self, df: pd.DataFrame, table_name: str) -> str:
        """Determine if table is Fact (metrics) or Dimension (attributes)"""
        
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        text_cols = df.select_dtypes(include=['object']).columns
        
        # Fact table: Mostly numeric columns (measures)
        if len(numeric_cols) >= len(df.columns) * 0.5:
            return 'Fact'
        
        # Dimension table: Mostly text/categorical
        return 'Dimension'
    
    def _detect_hierarchies(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Detect date/time hierarchies (Year > Quarter > Month > Day)"""
        hierarchies = []
        
        date_cols = df.select_dtypes(include=['datetime64']).columns
        
        for col in date_cols:
            hierarchy = {
                'name': f'{col}Hierarchy',
                'levels': [
                    {'name': 'Year', 'expression': f'YEAR({col})'},
                    {'name': 'Quarter', 'expression': f'QUARTER({col})'},
                    {'name': 'Month', 'expression': f'MONTH({col})'},
                    {'name': 'Day', 'expression': f'DAY({col})'}
                ]
            }
            hierarchies.append(hierarchy)
        
        return hierarchies
    
    def _auto_create_relationships(self):
        """Auto-detect relationships between tables (foreign keys)"""
        print(f"\n🔗 Detecting relationships...")
        
        tables = self.data_model['tables']
        
        for table1_name, table1_info in tables.items():
            df1 = table1_info['data']
            
            for table2_name, table2_info in tables.items():
                if table1_name == table2_name:
                    continue
                
                df2 = table2_info['data']
                
                # Look for matching column names
                for col1 in df1.columns:
                    for col2 in df2.columns:
                        # Check if column names suggest relationship
                        if self._is_potential_relationship(col1, col2, table1_name, table2_name):
                            # Verify data compatibility
                            if self._verify_relationship(df1[col1], df2[col2]):
                                relationship = {
                                    'from_table': table1_name,
                                    'from_column': col1,
                                    'to_table': table2_name,
                                    'to_column': col2,
                                    'cardinality': 'many-to-one',
                                    'cross_filter': 'single'
                                }
                                
                                # Avoid duplicates
                                if relationship not in self.relationships:
                                    self.relationships.append(relationship)
                                    print(f"   ✓ {table1_name}[{col1}] → {table2_name}[{col2}]")
    
    def _is_potential_relationship(self, col1: str, col2: str, 
                                   table1: str, table2: str) -> bool:
        """Check if column names suggest a foreign key relationship"""
        
        # Exact match
        if col1 == col2:
            return True
        
        # Foreign key pattern (e.g., ProductId in Sales table)
        if col1.lower().endswith('id') and table2.lower() in col1.lower():
            return True
        
        if col2.lower().endswith('id') and table1.lower() in col2.lower():
            return True
        
        return False
    
    def _verify_relationship(self, series1: pd.Series, series2: pd.Series) -> bool:
        """Verify that relationship is valid (values in series1 exist in series2)"""
        
        # Skip if too many missing values
        if series1.isna().sum() > len(series1) * 0.5:
            return False
        
        # Check if values in series1 exist in series2
        series1_unique = set(series1.dropna().unique())
        series2_unique = set(series2.dropna().unique())
        
        # At least 50% overlap
        overlap = len(series1_unique.intersection(series2_unique))
        return overlap >= len(series1_unique) * 0.5
    
    def _generate_dax_measures(self):
        """Auto-generate common DAX measures for each numeric column"""
        print(f"\n📊 Generating DAX measures...")
        
        for table_name, table_info in self.data_model['tables'].items():
            df = table_info['data']
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            
            for col in numeric_cols:
                # Skip ID columns
                if 'id' in col.lower():
                    continue
                
                # Total (SUM)
                self.dax_measures.append({
                    'name': f'Total{col}',
                    'table': table_name,
                    'expression': f'SUM({table_name}[{col}])',
                    'format': '0.00'
                })
                
                # Average
                self.dax_measures.append({
                    'name': f'Average{col}',
                    'table': table_name,
                    'expression': f'AVERAGE({table_name}[{col}])',
                    'format': '0.00'
                })
                
                # Count
                self.dax_measures.append({
                    'name': f'Count{col}',
                    'table': table_name,
                    'expression': f'COUNT({table_name}[{col}])',
                    'format': '0'
                })
                
                # Year-over-Year (if date column exists)
                date_cols = df.select_dtypes(include=['datetime64']).columns
                if len(date_cols) > 0:
                    date_col = date_cols[0]
                    self.dax_measures.append({
                        'name': f'{col}YoY',
                        'table': table_name,
                        'expression': f'CALCULATE(SUM({table_name}[{col}]), SAMEPERIODLASTYEAR({table_name}[{date_col}]))',
                        'format': '0.00%'
                    })
            
            table_info['measures'] = [m for m in self.dax_measures if m['table'] == table_name]
        
        print(f"   ✓ Generated {len(self.dax_measures)} DAX measures")
    
    def _package_model(self) -> Dict[str, Any]:
        """Package everything for Power BI"""
        
        return {
            'model': {
                'tables': {name: {
                    'name': info['name'],
                    'type': info['type'],
                    'columns': info['columns'],
                    'row_count': info['row_count'],
                    'hierarchies': info['hierarchies'],
                    'data': info['data'].to_dict('records')
                } for name, info in self.data_model['tables'].items()},
                'relationships': self.relationships,
                'measures': self.dax_measures
            },
            'metadata': {
                'tables_count': len(self.data_model['tables']),
                'relationships_count': len(self.relationships),
                'measures_count': len(self.dax_measures),
                'created_at': datetime.utcnow().isoformat()
            }
        }
    
    def export_to_csv(self, output_dir: str):
        """Export cleaned tables as CSV for Power BI import"""
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        for table_name, table_info in self.data_model['tables'].items():
            csv_path = output_path / f"{table_name}.csv"
            table_info['data'].to_csv(csv_path, index=False)
            print(f"   ✓ Exported: {csv_path}")


# Helper function
def convert_excel_to_powerbi_model(excel_path: str) -> Dict[str, Any]:
    """
    Quick function to convert Excel to Power BI model.
    
    Args:
        excel_path: Path to Excel file
        
    Returns:
        Power BI data model specification
    """
    processor = ExcelToPowerBIProcessor()
    return processor.process_excel(excel_path)
