"""
Template management API endpoints
Handles template listing, preview, and custom template management
"""

from fastapi import APIRouter, HTTPException, UploadFile, File, Depends, status
from fastapi.responses import FileResponse, JSONResponse
from typing import List, Optional
import os
import sys
import json
import tempfile
from pathlib import Path

# Add paths for imports
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../../../../../')))

from src.templates.template_manager import TemplateManager, get_default_template
from api.deps import get_current_user
from models.user import UserInDB as User

router = APIRouter()

# Initialize template manager
template_manager = TemplateManager()


@router.get("/templates", response_model=List[dict])
async def list_templates(
    category: Optional[str] = None,
    include_custom: bool = True
):
    """
    List all available templates
    
    Args:
        category: Filter by category (business, finance, technology, etc.)
        include_custom: Include custom user templates
    
    Returns:
        List of template information
    """
    try:
        templates = template_manager.list_templates(include_custom=include_custom)
        
        # Filter by category if specified
        if category:
            templates = [t for t in templates if t.get('category') == category]
        
        return templates
    
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error listing templates: {str(e)}"
        )


@router.get("/templates/{template_name}")
async def get_template(template_name: str):
    """
    Get detailed template configuration
    
    Args:
        template_name: Name of the template
    
    Returns:
        Template configuration JSON
    """
    try:
        template = template_manager.load_template(template_name)
        
        if not template:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Template '{template_name}' not found"
            )
        
        return template
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error loading template: {str(e)}"
        )


@router.get("/templates/{template_name}/colors")
async def get_template_colors(template_name: str):
    """
    Get color palette for a template
    
    Args:
        template_name: Name of the template
    
    Returns:
        Color configuration
    """
    try:
        template = template_manager.load_template(template_name)
        
        if not template:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Template '{template_name}' not found"
            )
        
        return {
            "template": template_name,
            "colors": template.get("colors", {}),
            "chart_colors": template.get("colors", {}).get("chart_colors", [])
        }
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error getting template colors: {str(e)}"
        )


@router.post("/templates/custom")
async def create_custom_template(
    file: UploadFile = File(...),
    name: Optional[str] = None,
    current_user: User = Depends(get_current_user)
):
    """
    Upload a custom template JSON file
    
    Args:
        file: Template JSON file
        name: Optional custom name for the template
        current_user: Authenticated user
    
    Returns:
        Success message with template name
    """
    try:
        # Read uploaded file
        content = await file.read()
        template_data = json.loads(content)
        
        # Validate template
        if not template_manager.validate_template(template_data):
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Invalid template format. Please check required fields."
            )
        
        # Use provided name or filename
        template_name = name or Path(file.filename).stem
        
        # Save custom template
        success = template_manager.save_custom_template(
            template_data,
            template_name,
            overwrite=False
        )
        
        if not success:
            raise HTTPException(
                status_code=status.HTTP_409_CONFLICT,
                detail=f"Template '{template_name}' already exists"
            )
        
        return {
            "message": "Template created successfully",
            "template_name": template_name,
            "template_id": template_name
        }
    
    except json.JSONDecodeError:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Invalid JSON file"
        )
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error creating template: {str(e)}"
        )


@router.delete("/templates/custom/{template_name}")
async def delete_custom_template(
    template_name: str,
    current_user: User = Depends(get_current_user)
):
    """
    Delete a custom template (built-in templates cannot be deleted)
    
    Args:
        template_name: Name of the template to delete
        current_user: Authenticated user
    
    Returns:
        Success message
    """
    try:
        # Check if template exists and is custom
        template = template_manager.load_template(template_name)
        
        if not template:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Template '{template_name}' not found"
            )
        
        if template.get('source') == 'built-in':
            raise HTTPException(
                status_code=status.HTTP_403_FORBIDDEN,
                detail="Cannot delete built-in templates"
            )
        
        # Delete template
        success = template_manager.delete_custom_template(template_name)
        
        if not success:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="Failed to delete template"
            )
        
        return {
            "message": f"Template '{template_name}' deleted successfully"
        }
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error deleting template: {str(e)}"
        )


@router.post("/templates/{template_name}/duplicate")
async def duplicate_template(
    template_name: str,
    new_name: str,
    current_user: User = Depends(get_current_user)
):
    """
    Duplicate a template to create a custom version
    
    Args:
        template_name: Source template name
        new_name: Name for the new template
        current_user: Authenticated user
    
    Returns:
        Success message with new template name
    """
    try:
        success = template_manager.duplicate_template(template_name, new_name)
        
        if not success:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Failed to duplicate template. Source template may not exist."
            )
        
        return {
            "message": "Template duplicated successfully",
            "template_name": new_name,
            "source_template": template_name
        }
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error duplicating template: {str(e)}"
        )


@router.get("/templates/{template_name}/export")
async def export_template(
    template_name: str,
    current_user: User = Depends(get_current_user)
):
    """
    Export a template as a JSON file
    
    Args:
        template_name: Template to export
        current_user: Authenticated user
    
    Returns:
        JSON file download
    """
    try:
        template = template_manager.load_template(template_name)
        
        if not template:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=f"Template '{template_name}' not found"
            )
        
        # Remove internal metadata
        template.pop('source', None)
        template.pop('path', None)
        template.pop('created_at', None)
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(template, f, indent=2)
            temp_path = f.name
        
        return FileResponse(
            temp_path,
            media_type='application/json',
            filename=f"{template_name}.json"
        )
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error exporting template: {str(e)}"
        )


@router.get("/templates/default")
async def get_default():
    """Get the default template name"""
    return {
        "default_template": get_default_template(),
        "description": "Default template used when none is specified"
    }


@router.get("/templates/categories")
async def get_categories():
    """Get list of template categories"""
    templates = template_manager.list_templates()
    categories = list(set(t.get('category', 'general') for t in templates))
    
    return {
        "categories": sorted(categories),
        "count": len(categories)
    }
