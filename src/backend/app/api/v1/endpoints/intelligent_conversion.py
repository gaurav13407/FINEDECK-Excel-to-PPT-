"""
API Integration for Data Intelligence Engine
Integrates the automatic data analysis into FinDeck's conversion flow.
"""

from fastapi import APIRouter, UploadFile, File, HTTPException, Depends
from fastapi.responses import FileResponse
from typing import Dict, Any
import tempfile
import os
import sys
from pathlib import Path

# Add path to find converter modules
converter_path = os.path.join(os.path.dirname(__file__), "../../../../../")
if converter_path not in sys.path:
    sys.path.insert(0, converter_path)

from converter.excel_reader import excel_reader
from converter.ppt_writer import df_to_ppt
from services.data_intelligence import DataIntelligenceEngine
from api.deps import get_current_active_user
from models.user import UserInDB

router = APIRouter()


@router.post("/analyze-excel", response_model=Dict[str, Any])
async def analyze_excel_file(
    file: UploadFile = File(...),
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Analyze uploaded Excel file and return structured insights.
    
    This endpoint:
    1. Accepts any Excel file
    2. Automatically classifies all columns
    3. Extracts metrics for each column type
    4. Identifies hierarchy (Sales, Profit, Geo)
    5. Returns PPT-ready JSON output
    
    Zero hallucination - 100% based on actual data.
    """
    
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Only Excel files (.xlsx, .xls) are supported."
        )
    
    # Save uploaded file temporarily
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            content = await file.read()
            temp_file.write(content)
            temp_path = temp_file.name
        
        # Analyze the file
        engine = DataIntelligenceEngine()
        analysis_results = engine.analyze_file(temp_path)
        
        # Clean up temp file
        os.unlink(temp_path)
        
        # Add metadata
        analysis_results['metadata'] = {
            'filename': file.filename,
            'user_id': str(current_user.id),
            'analyzed_at': engine.df.shape if hasattr(engine, 'df') else None
        }
        
        return {
            'success': True,
            'data': analysis_results,
            'message': 'File analyzed successfully'
        }
        
    except Exception as e:
        # Clean up on error
        if 'temp_path' in locals() and os.path.exists(temp_path):
            os.unlink(temp_path)
        
        raise HTTPException(
            status_code=500,
            detail=f"Error analyzing file: {str(e)}"
        )


@router.post("/convert-with-intelligence")
async def convert_excel_to_ppt_intelligent(
    file: UploadFile = File(...),
    title: str = "Intelligent Data Analysis",
    subtitle: str = "Auto-Generated Insights",
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Convert Excel to PPT using intelligent data analysis.
    
    This endpoint:
    1. Analyzes the Excel file automatically
    2. Generates insights and metrics  
    3. Creates PPT slides using your existing df_to_ppt converter
    4. Returns the PPT file for download
    
    This is the SMART conversion - no manual configuration needed!
    """
    
    # Validate file type
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(
            status_code=400,
            detail="Invalid file type. Only Excel files (.xlsx, .xls) are supported."
        )
    
    temp_excel_path = None
    temp_ppt_path = None
    
    try:
        # Save uploaded Excel file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            content = await file.read()
            temp_file.write(content)
            temp_excel_path = temp_file.name
        
        # STEP 1: Analyze the data with Intelligence Engine
        engine = DataIntelligenceEngine()
        analysis = engine.analyze_file(temp_excel_path)
        
        # STEP 2: Read Excel data using your existing reader
        df = excel_reader(temp_excel_path, sheet=0)
        
        if df is None or df.empty:
            raise HTTPException(
                status_code=400,
                detail="No data found in Excel file"
            )
        
        # STEP 3: Create PPT using your existing converter
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_ppt:
            temp_ppt_path = tmp_ppt.name
        
        # Generate intelligent title with data insights
        intelligent_title = f"{title} - {analysis['executive_summary']['total_records']:,} Records Analyzed"
        intelligent_subtitle = f"{subtitle} | {analysis['executive_summary']['numeric_columns']} Metrics | {analysis['executive_summary']['categorical_columns']} Categories"
        
        # Convert to PowerPoint using your existing df_to_ppt function
        # Limit to 50 rows for reasonable PPT size (can be made configurable)
        df_to_ppt(
            df=df,
            out_path=temp_ppt_path,
            title=intelligent_title,
            subtitle=intelligent_subtitle,
            title_col=None,  # Let it auto-detect
            mode="table",    # Use table mode for better visualization
            limit=50         # Limit to 50 rows for performance
        )
        
        # Generate filename for download
        base_name = Path(file.filename).stem
        ppt_filename = f"{base_name}_intelligent_analysis.pptx"
        
        # Return the PPT file
        return FileResponse(
            path=temp_ppt_path,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=ppt_filename,
            headers={
                "X-Analysis-Records": str(analysis['executive_summary']['total_records']),
                "X-Analysis-Columns": str(analysis['executive_summary']['total_columns']),
                "X-Numeric-Metrics": str(analysis['executive_summary']['numeric_columns'])
            }
        )
        
    except Exception as e:
        # Clean up temp files on error
        if temp_excel_path and os.path.exists(temp_excel_path):
            os.unlink(temp_excel_path)
        if temp_ppt_path and os.path.exists(temp_ppt_path):
            os.unlink(temp_ppt_path)
        
        raise HTTPException(
            status_code=500,
            detail=f"Error in intelligent conversion: {str(e)}"
        )
    finally:
        # Clean up Excel file (PPT will be cleaned up after response)
        if temp_excel_path and os.path.exists(temp_excel_path):
            os.unlink(temp_excel_path)


@router.get("/column-classification-rules")
async def get_classification_rules():
    """
    Get the rules used for automatic column classification.
    Useful for documentation and transparency.
    """
    return {
        'identifier_patterns': DataIntelligenceEngine.IDENTIFIER_PATTERNS,
        'sales_patterns': DataIntelligenceEngine.SALES_PATTERNS,
        'profit_patterns': DataIntelligenceEngine.PROFIT_PATTERNS,
        'geo_patterns': DataIntelligenceEngine.GEO_PATTERNS,
        'column_types': [
            'numeric',
            'categorical',
            'text',
            'date',
            'boolean',
            'identifier'
        ],
        'rules': {
            'identifiers': 'Never summed - treated as unique IDs',
            'numeric': 'Calculate sum, mean, median, min, max, std, top 10',
            'categorical': 'Count unique, top 10 categories, percentage distribution',
            'date': 'Monthly/yearly grouping, trend detection',
            'boolean': 'True/false count and percentages'
        }
    }
