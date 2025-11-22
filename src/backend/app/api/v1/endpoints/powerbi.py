"""
Power BI Dashboard API Endpoints
Automatically converts Excel to Power BI dashboards
"""

from fastapi import APIRouter, UploadFile, File, HTTPException, Depends, Form
from fastapi.responses import JSONResponse, FileResponse
from typing import Optional, List, Dict, Any
from pathlib import Path
import shutil
import os
from datetime import datetime

from src.PoweBI_converter.powerbi_etl import ExcelToPowerBIProcessor
from src.PoweBI_converter.powerbi_export import export_simple_package
from app.models.user import UserInDB
from app.api.deps import get_current_active_user
from app.core.config import settings

router = APIRouter()


# Power BI Dashboard Templates
DASHBOARD_TEMPLATES = {
    'revenue_profit': {
        'name': 'Revenue & Profit Dashboard',
        'description': 'Track revenue streams, profit margins, and financial trends',
        'visuals': ['Line Chart (Revenue Trend)', 'Bar Chart (Profit by Category)', 
                   'KPI Cards (Total Revenue, Profit Margin)', 'Donut Chart (Revenue Mix)'],
        'data_requirements': ['Revenue', 'Cost', 'Date', 'Category']
    },
    'sales_performance': {
        'name': 'Sales Performance Dashboard',
        'description': 'Analyze sales by region, product, and sales rep',
        'visuals': ['Map (Sales by Region)', 'Table (Top Products)', 
                   'Column Chart (Monthly Sales)', 'Gauge (Target Achievement)'],
        'data_requirements': ['Sales', 'Quantity', 'Region', 'Product', 'Date']
    },
    'financial_kpi': {
        'name': 'Financial KPI Dashboard',
        'description': 'Monitor key financial metrics and performance indicators',
        'visuals': ['KPI Cards (Revenue, Profit, Expenses, Margin)', 
                   'Waterfall Chart (P&L)', 'Area Chart (Cash Flow)', 'Gauge (Budget vs Actual)'],
        'data_requirements': ['Revenue', 'Expenses', 'Profit', 'Budget', 'Date']
    },
    'marketing_analytics': {
        'name': 'Marketing Analytics Dashboard',
        'description': 'Track campaigns, conversions, and customer acquisition',
        'visuals': ['Funnel Chart (Conversion)', 'Line Chart (CAC Trend)', 
                   'Bar Chart (Campaign ROI)', 'Scatter Plot (Engagement vs Spend)'],
        'data_requirements': ['Impressions', 'Clicks', 'Conversions', 'Spend', 'Campaign', 'Date']
    },
    'operations_efficiency': {
        'name': 'Operations Efficiency Dashboard',
        'description': 'Monitor operational metrics, productivity, and resource utilization',
        'visuals': ['Column Chart (Production Volume)', 'Line Chart (Efficiency Trend)', 
                   'KPI Cards (Utilization, Downtime)', 'Heatmap (Resource Allocation)'],
        'data_requirements': ['Production', 'Capacity', 'Downtime', 'Resource', 'Date']
    }
}


@router.post("/create-dashboard")
async def create_powerbi_dashboard(
    file: UploadFile = File(...),
    template: Optional[str] = Form('auto'),
    dashboard_title: Optional[str] = Form(None),
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Create Power BI dashboard from Excel file.
    
    Process:
    1. Upload Excel file
    2. Run ETL pipeline (clean, normalize, model)
    3. Detect data type or use selected template
    4. Generate Power BI model specification
    5. Return dashboard metadata + download link
    
    Args:
        file: Excel file (.xlsx, .xls)
        template: Dashboard template ('auto', 'revenue_profit', 'sales_performance', etc.)
        dashboard_title: Custom dashboard title
        current_user: Authenticated user
        
    Returns:
        Dashboard metadata with data model and measures
    """
    
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are supported")
    
    # Create temp directory for processing
    temp_dir = Path(settings.TEMP_DIR) / f"powerbi_{current_user.id}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
    temp_dir.mkdir(parents=True, exist_ok=True)
    
    try:
        # Save uploaded file
        excel_path = temp_dir / file.filename
        with open(excel_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        print(f"📁 Processing Excel file: {file.filename}")
        
        # Step 1: Run ETL pipeline
        processor = ExcelToPowerBIProcessor()
        powerbi_model = processor.process_excel(str(excel_path))
        
        # Step 2: Auto-detect template if not specified
        if template == 'auto':
            detected_template = _detect_template(powerbi_model)
            template = detected_template
            print(f"🔍 Auto-detected template: {DASHBOARD_TEMPLATES[template]['name']}")
        
        # Validate template
        if template not in DASHBOARD_TEMPLATES:
            raise HTTPException(status_code=400, 
                              detail=f"Invalid template. Choose from: {list(DASHBOARD_TEMPLATES.keys())}")
        
        # Step 3: Apply template metadata
        template_info = DASHBOARD_TEMPLATES[template]
        
        # Step 4: Generate dashboard title
        if not dashboard_title:
            dashboard_title = f"{file.filename.split('.')[0]} - {template_info['name']}"
        
        # Step 5: Create simple ZIP package (CSV + README + DAX measures)
        package_file = export_simple_package(str(excel_path), output_dir=str(temp_dir))
        
        # Step 6: Simple response
        response = {
            'dashboard': {
                'id': f"pbi_{current_user.id}_{int(datetime.now().timestamp())}",
                'title': dashboard_title,
                'template': template,
                'template_name': template_info['name'],
                'description': template_info['description'],
                'created_at': datetime.utcnow().isoformat(),
                'owner_id': current_user.id,
                'owner_email': current_user.email
            },
            'data_model': {
                'tables_count': powerbi_model['metadata']['tables_count'],
                'measures_count': powerbi_model['metadata']['measures_count'],
                'relationships_count': powerbi_model['metadata']['relationships_count']
            },
            'download': {
                'file': str(package_file),
                'filename': package_file.name,
                'size_kb': round(package_file.stat().st_size / 1024, 2),
                'type': 'ZIP Package'
            },
            'next_steps': [
                "1. Download the ZIP file",
                "2. Extract the files",
                "3. Open Power BI Desktop",
                "4. Import CSV files from data/ folder",
                "5. Copy DAX measures from DAX_Measures.txt"
            ]
        }
        
        print(f"✅ Dashboard created: {dashboard_title}")
        print(f"   📊 Tables: {powerbi_model['metadata']['tables_count']}")
        print(f"   🔗 Relationships: {powerbi_model['metadata']['relationships_count']}")
        print(f"   📈 Measures: {powerbi_model['metadata']['measures_count']}")
        print(f"   📦 Package: {package_file.name} ({response['download']['size_kb']} KB)")
        
        return JSONResponse(content=response, status_code=201)
        
    except Exception as e:
        print(f"❌ Error creating dashboard: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Dashboard creation failed: {str(e)}")
    
    finally:
        # Cleanup temp files (optional - keep for debugging)
        # shutil.rmtree(temp_dir, ignore_errors=True)
        pass


@router.get("/templates")
async def get_dashboard_templates():
    """
    Get all available Power BI dashboard templates.
    
    Returns:
        List of templates with metadata
    """
    return {
        'templates': [
            {
                'id': template_id,
                **template_info
            }
            for template_id, template_info in DASHBOARD_TEMPLATES.items()
        ],
        'count': len(DASHBOARD_TEMPLATES)
    }


@router.get("/dashboards")
async def list_user_dashboards(
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    List all Power BI dashboards created by the user.
    
    TODO: Implement MongoDB storage for dashboard metadata
    """
    # Placeholder - will implement MongoDB storage
    return {
        'dashboards': [],
        'count': 0,
        'message': 'Dashboard storage coming soon'
    }


@router.get("/status/{dashboard_id}")
async def get_dashboard_status(
    dashboard_id: str,
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Get status of a Power BI dashboard.
    
    TODO: Implement dashboard status tracking
    """
    return {
        'dashboard_id': dashboard_id,
        'status': 'ready',
        'message': 'Dashboard status tracking coming soon'
    }


@router.get("/download/{dashboard_id}")
async def download_dashboard_file(
    dashboard_id: str,
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Download Power BI package (ZIP with CSV + README + DAX measures).
    
    Args:
        dashboard_id: Dashboard ID (e.g., pbi_user123_1234567890)
        current_user: Authenticated user
        
    Returns:
        ZIP file download
    """
    
    # In production, retrieve file path from MongoDB
    # For now, find most recent ZIP file in temp directory
    temp_dir = Path(settings.TEMP_DIR)
    
    # Search for ZIP files
    zip_files = list(temp_dir.glob(f"powerbi_*/*_powerbi_package.zip"))
    
    if not zip_files:
        raise HTTPException(status_code=404, detail="Dashboard file not found")
    
    # Get most recent file
    zip_file = max(zip_files, key=lambda p: p.stat().st_mtime)
    
    if not zip_file.exists():
        raise HTTPException(status_code=404, detail="File not found")
    
    return FileResponse(
        path=str(zip_file),
        media_type='application/zip',
        filename=f"powerbi_dashboard_{dashboard_id}.zip"
    )


def _detect_template(powerbi_model: Dict[str, Any]) -> str:
    """
    Auto-detect appropriate dashboard template based on data.
    
    Logic:
    - If has Revenue/Profit columns → revenue_profit
    - If has Sales/Quantity columns → sales_performance
    - If has Budget/Expenses columns → financial_kpi
    - If has Campaign/Clicks columns → marketing_analytics
    - Default → operations_efficiency
    """
    
    # Get all column names from all tables
    all_columns = []
    for table_name, table_info in powerbi_model['model']['tables'].items():
        all_columns.extend([col.lower() for col in table_info['columns']])
    
    columns_text = ' '.join(all_columns)
    
    # Revenue & Profit
    if any(keyword in columns_text for keyword in ['revenue', 'profit', 'margin', 'income']):
        return 'revenue_profit'
    
    # Sales Performance
    if any(keyword in columns_text for keyword in ['sales', 'quantity', 'region', 'product']):
        return 'sales_performance'
    
    # Financial KPI
    if any(keyword in columns_text for keyword in ['budget', 'expenses', 'cost', 'pl', 'cashflow']):
        return 'financial_kpi'
    
    # Marketing Analytics
    if any(keyword in columns_text for keyword in ['campaign', 'clicks', 'impressions', 'conversions', 'cac']):
        return 'marketing_analytics'
    
    # Default: Operations
    return 'operations_efficiency'
