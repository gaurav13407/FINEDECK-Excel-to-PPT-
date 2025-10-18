# Excel to PowerPoint Conversion Endpoints
# Basic conversion functionality using existing ppt_writer
# - POST /conversions/convert - Direct Excel to PPT conversion
# - GET /conversions/templates - Get available PowerPoint templates

from fastapi import APIRouter, Depends, HTTPException, status, BackgroundTasks, Form
from fastapi.responses import FileResponse
from typing import List, Optional, Dict, Any
from datetime import datetime
import os
import sys
import tempfile
from pathlib import Path

# Add path to find converter modules
converter_path = os.path.join(os.path.dirname(__file__), "../../../../../")
if converter_path not in sys.path:
    sys.path.insert(0, converter_path)

# Import your existing converters
from converter.excel_reader import excel_reader
from converter.ppt_writer import df_to_ppt
from api.deps import get_current_active_user, require_credits
from models.user import UserInDB
from services.file_service import get_file_by_id

router = APIRouter()

@router.post("/convert")
async def convert_excel_to_ppt(
    file_id: str = Form(...),
    title: str = Form("Auto Report"),
    subtitle: str = Form(""),
    sheet_name: Optional[str] = Form(None),
    title_col: Optional[str] = Form(None),
    mode: str = Form("table"),  # "table" or "text"
    limit: Optional[int] = Form(None),
    current_user: UserInDB = Depends(get_current_active_user),
    _: None = Depends(require_credits(required_credits=1))
):
    """
    Convert an Excel file to PowerPoint presentation
    
    Args:
        file_id: ID of the uploaded Excel file
        title: Presentation title
        subtitle: Presentation subtitle  
        sheet_name: Specific sheet name to convert (optional)
        title_col: Column to use for slide titles (optional)
        mode: "table" or "text" - how to display data
        limit: Maximum number of rows to convert (optional)
    """
    try:
        # Get the uploaded file
        file_doc = await get_file_by_id(file_id)
        if not file_doc:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found"
            )
        
        # Check if user owns the file
        if file_doc.user_id != current_user.id:
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Access denied to this file"
            )
        
        # Check if it's an Excel file
        if not file_doc.filename.lower().endswith(('.xlsx', '.xls')):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="File must be an Excel file (.xlsx or .xls)"
            )
        
        # Read Excel file
        excel_path = file_doc.storage_path  # Assuming this is the local path
        if not os.path.exists(excel_path):
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="Excel file not found on storage"
            )
        
        # Convert sheet_name to appropriate type
        sheet_param = None
        if sheet_name:
            # Try to convert to int if it's numeric, otherwise keep as string
            try:
                sheet_param = int(sheet_name)
            except ValueError:
                sheet_param = sheet_name
        
        # Read the Excel data
        df = excel_reader(excel_path, sheet=sheet_param)
        
        if df is None or df.empty:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="No data found in Excel file"
            )
        
        # Create temporary PPT file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp_file:
            ppt_path = tmp_file.name
        
        # Convert to PowerPoint
        df_to_ppt(
            df=df,
            out_path=ppt_path,
            title=title,
            subtitle=subtitle,
            title_col=title_col,
            mode=mode,
            limit=limit
        )
        
        # Generate filename for download
        base_name = Path(file_doc.filename).stem
        ppt_filename = f"{base_name}_converted.pptx"
        
        # Return the file for download
        return FileResponse(
            path=ppt_path,
            filename=ppt_filename,
            media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )
        
    except Exception as e:
        # Clean up temp file if it exists
        if 'ppt_path' in locals() and os.path.exists(ppt_path):
            os.unlink(ppt_path)
        
        if isinstance(e, HTTPException):
            raise e
        
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Conversion failed: {str(e)}"
        )

@router.get("/templates")
async def get_available_templates(
    current_user: UserInDB = Depends(get_current_active_user)
):
    """
    Get list of available PowerPoint templates
    For now, returns basic template info
    """
    return {
        "templates": [
            {
                "id": "default",
                "name": "Default Template", 
                "description": "Basic PowerPoint template with title and content slides"
            },
            {
                "id": "table",
                "name": "Table Layout",
                "description": "Optimized for displaying data in table format"
            },
            {
                "id": "text", 
                "name": "Text Layout",
                "description": "Optimized for bullet point text content"
            }
        ]
    }