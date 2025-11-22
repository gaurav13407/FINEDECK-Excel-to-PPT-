"""
Automated Excel → PowerBI + PPT Pipeline
ONE UPLOAD, TWO OUTPUTS IN SECONDS

This service automates 80% of user work by:
1. Auto-detecting data type from Excel
2. Generating PowerBI dashboard automatically
3. Creating matching PPT presentation
4. Returning both files ready to use

No manual work required - just upload and download.
"""

import pandas as pd
from pathlib import Path
from typing import Dict, Any, Optional, Tuple, List
from datetime import datetime
import asyncio
from concurrent.futures import ThreadPoolExecutor

from src.PoweBI_converter.powerbi_etl import ExcelToPowerBIProcessor
from src.PoweBI_converter.powerbi_export import export_simple_package
from app.services.smart_ppt_service import SmartPPTGenerator


class AutomatedPipelineService:
    """
    Fully automated Excel → PowerBI + PPT pipeline.
    User uploads once, gets both outputs in seconds.
    """
    
    def __init__(self):
        self.powerbi_processor = ExcelToPowerBIProcessor()
        self.ppt_generator = SmartPPTGenerator()
        self.executor = ThreadPoolExecutor(max_workers=2)
        
    async def process_excel_complete(
        self,
        excel_path: str,
        user_preferences: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Main automated pipeline: Excel → PowerBI + PPT
        
        Args:
            excel_path: Path to uploaded Excel file
            user_preferences: Optional customization (template, colors, etc.)
            
        Returns:
            {
                'powerbi': {
                    'file_path': str,
                    'dashboard_type': str,
                    'tables': int,
                    'measures': int,
                    'download_url': str
                },
                'powerpoint': {
                    'file_path': str,
                    'slides': int,
                    'charts': int,
                    'download_url': str
                },
                'metadata': {
                    'processing_time': float,
                    'data_rows': int,
                    'data_columns': int,
                    'detected_type': str
                }
            }
        """
        
        start_time = datetime.now()
        print(f"🚀 Starting Automated Pipeline: {Path(excel_path).name}")
        
        # Step 1: Analyze Excel to detect data type
        data_analysis = self._analyze_excel(excel_path)
        
        # Step 2: Run PowerBI and PPT generation in parallel
        powerbi_result, ppt_result = await self._parallel_generation(
            excel_path=excel_path,
            data_analysis=data_analysis,
            user_preferences=user_preferences or {}
        )
        
        # Step 3: Package results
        processing_time = (datetime.now() - start_time).total_seconds()
        
        result = {
            'powerbi': powerbi_result,
            'powerpoint': ppt_result,
            'metadata': {
                'processing_time': processing_time,
                'data_rows': data_analysis['total_rows'],
                'data_columns': data_analysis['total_columns'],
                'detected_type': data_analysis['data_type'],
                'timestamp': datetime.now().isoformat()
            }
        }
        
        print(f"✅ Pipeline Complete in {processing_time:.2f}s")
        print(f"   📊 PowerBI: {powerbi_result['dashboard_type']}")
        print(f"   📄 PPT: {ppt_result['slides']} slides, {ppt_result['charts']} charts")
        
        return result
    
    def _analyze_excel(self, excel_path: str) -> Dict[str, Any]:
        """
        Automatically detect what type of data this is.
        
        Detects:
        - Financial data (Revenue, Profit, Expenses)
        - Sales data (Quantity, Price, Region)
        - Marketing data (Impressions, Clicks, Conversions)
        - Operations data (Production, Capacity, Efficiency)
        - Custom/Other
        
        Returns:
            Analysis with detected type and recommendations
        """
        
        print("🔍 Analyzing Excel data...")
        
        # Load all sheets
        excel_file = pd.ExcelFile(excel_path)
        all_data = {}
        total_rows = 0
        total_columns = 0
        all_column_names = []
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(excel_path, sheet_name=sheet_name)
            if not df.empty:
                all_data[sheet_name] = df
                total_rows += len(df)
                total_columns += len(df.columns)
                all_column_names.extend([col.lower() for col in df.columns])
        
        # Detect data type by analyzing column names
        data_type = self._detect_data_type(all_column_names)
        
        # Get recommended templates
        powerbi_template = self._get_powerbi_template(data_type)
        ppt_template = self._get_ppt_template(data_type)
        
        analysis = {
            'total_rows': total_rows,
            'total_columns': total_columns,
            'sheets': list(all_data.keys()),
            'data_type': data_type,
            'powerbi_template': powerbi_template,
            'ppt_template': ppt_template,
            'column_names': all_column_names
        }
        
        print(f"   ✓ Detected: {data_type}")
        print(f"   ✓ PowerBI Template: {powerbi_template}")
        print(f"   ✓ PPT Template: {ppt_template}")
        
        return analysis
    
    def _detect_data_type(self, column_names: List[str]) -> str:
        """
        Detect data type based on column names.
        Uses keyword matching to identify the domain.
        """
        
        # Convert all to lowercase for matching
        columns = ' '.join(column_names).lower()
        
        # Financial keywords
        financial_keywords = ['revenue', 'profit', 'expense', 'cost', 'margin', 
                            'ebitda', 'cash', 'balance', 'income', 'sales']
        financial_score = sum(1 for kw in financial_keywords if kw in columns)
        
        # Sales keywords
        sales_keywords = ['quantity', 'price', 'product', 'region', 'sales', 
                         'customer', 'order', 'unit']
        sales_score = sum(1 for kw in sales_keywords if kw in columns)
        
        # Marketing keywords
        marketing_keywords = ['impression', 'click', 'conversion', 'ctr', 'cac', 
                            'campaign', 'ad', 'spend', 'roi']
        marketing_score = sum(1 for kw in marketing_keywords if kw in columns)
        
        # Operations keywords
        operations_keywords = ['production', 'capacity', 'efficiency', 'downtime', 
                             'utilization', 'resource', 'output']
        operations_score = sum(1 for kw in operations_keywords if kw in columns)
        
        # Determine type based on highest score
        scores = {
            'financial': financial_score,
            'sales': sales_score,
            'marketing': marketing_score,
            'operations': operations_score
        }
        
        max_score = max(scores.values())
        if max_score == 0:
            return 'general'
        
        return max(scores, key=scores.get)
    
    def _get_powerbi_template(self, data_type: str) -> str:
        """Map data type to PowerBI dashboard template"""
        
        template_map = {
            'financial': 'financial_kpi',
            'sales': 'sales_performance',
            'marketing': 'marketing_analytics',
            'operations': 'operations_efficiency',
            'general': 'revenue_profit'  # Default fallback
        }
        
        return template_map.get(data_type, 'revenue_profit')
    
    def _get_ppt_template(self, data_type: str) -> str:
        """Map data type to PPT template"""
        
        template_map = {
            'financial': 'financial_report',
            'sales': 'sales_dashboard',
            'marketing': 'marketing_report',
            'operations': 'operations_report',
            'general': 'modern_corporate'
        }
        
        return template_map.get(data_type, 'modern_corporate')
    
    async def _parallel_generation(
        self,
        excel_path: str,
        data_analysis: Dict[str, Any],
        user_preferences: Dict[str, Any]
    ) -> Tuple[Dict[str, Any], Dict[str, Any]]:
        """
        Generate PowerBI and PPT in parallel for maximum speed.
        Uses asyncio + ThreadPoolExecutor for true parallelism.
        """
        
        print("⚡ Running parallel generation (PowerBI + PPT)...")
        
        # Create event loop tasks
        loop = asyncio.get_event_loop()
        
        # Task 1: Generate PowerBI dashboard
        powerbi_task = loop.run_in_executor(
            self.executor,
            self._generate_powerbi,
            excel_path,
            data_analysis['powerbi_template'],
            user_preferences
        )
        
        # Task 2: Generate PPT presentation
        ppt_task = loop.run_in_executor(
            self.executor,
            self._generate_ppt,
            excel_path,
            data_analysis['ppt_template'],
            user_preferences
        )
        
        # Wait for both to complete
        powerbi_result, ppt_result = await asyncio.gather(powerbi_task, ppt_task)
        
        return powerbi_result, ppt_result
    
    def _generate_powerbi(
        self,
        excel_path: str,
        template: str,
        preferences: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Generate PowerBI dashboard file.
        Returns file path and metadata.
        """
        
        print("📊 Generating PowerBI dashboard...")
        
        try:
            # Process Excel through ETL pipeline
            dashboard_model = self.powerbi_processor.process_excel(excel_path)
            
            # Export to PowerBI package
            output_dir = Path(excel_path).parent / "powerbi_output"
            output_dir.mkdir(exist_ok=True)
            
            output_file = export_simple_package(
                excel_path=excel_path,
                output_dir=str(output_dir)
            )
            
            return {
                'file_path': str(output_file),
                'dashboard_type': template,
                'tables': len(dashboard_model['model']['tables']),
                'measures': len(dashboard_model['model']['measures']),
                'relationships': len(dashboard_model['model']['relationships']),
                'download_url': f"/downloads/powerbi/{output_file.name}",
                'status': 'success'
            }
            
        except Exception as e:
            print(f"❌ PowerBI generation failed: {e}")
            return {
                'status': 'error',
                'error': str(e),
                'file_path': None
            }
    
    def _generate_ppt(
        self,
        excel_path: str,
        template: str,
        preferences: Dict[str, Any]
    ) -> Dict[str, Any]:
        """
        Generate PowerPoint presentation.
        Returns file path and metadata.
        """
        
        print("📄 Generating PowerPoint presentation...")
        
        try:
            # Use existing tiered conversion service
            output_dir = Path(excel_path).parent / "ppt_output"
            output_dir.mkdir(exist_ok=True)
            
            output_file = output_dir / f"{Path(excel_path).stem}_presentation.pptx"
            
            # Generate PPT using smart generator
            ppt_result = self.ppt_generator.create_presentation(
                excel_path=excel_path,
                output_path=str(output_file),
                template_style=template
            )
            
            return {
                'file_path': str(output_file),
                'slides': ppt_result.get('total_slides', 0),
                'charts': ppt_result.get('charts_created', 0),
                'template_used': template,
                'download_url': f"/downloads/ppt/{output_file.name}",
                'status': 'success'
            }
            
        except Exception as e:
            print(f"❌ PPT generation failed: {e}")
            return {
                'status': 'error',
                'error': str(e),
                'file_path': None
            }
    
    def get_processing_estimate(self, excel_path: str) -> Dict[str, Any]:
        """
        Quick preview: What will be generated?
        Returns estimate without actually processing.
        """
        
        analysis = self._analyze_excel(excel_path)
        
        return {
            'data_type': analysis['data_type'],
            'powerbi_dashboard': analysis['powerbi_template'],
            'ppt_template': analysis['ppt_template'],
            'estimated_time': '5-15 seconds',
            'outputs': [
                'PowerBI Dashboard Package (.zip)',
                'PowerPoint Presentation (.pptx)'
            ],
            'data_summary': {
                'rows': analysis['total_rows'],
                'columns': analysis['total_columns'],
                'sheets': len(analysis['sheets'])
            }
        }


# Convenience function for quick access
async def auto_convert_excel(
    excel_path: str,
    preferences: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    One-line function to convert Excel to PowerBI + PPT.
    
    Usage:
        result = await auto_convert_excel("data.xlsx")
        print(result['powerbi']['download_url'])
        print(result['powerpoint']['download_url'])
    """
    
    service = AutomatedPipelineService()
    return await service.process_excel_complete(excel_path, preferences)
