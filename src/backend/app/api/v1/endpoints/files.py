# File Management Endpoints
# Handles Excel file uploads and management
# - POST /files/upload - Upload Excel files with validation
# - GET /files - List user's uploaded files with pagination
# - GET /files/{file_id} - Get specific file metadata
# - DELETE /files/{file_id} - Delete uploaded file
# - GET /files/{file_id}/download - Download original Excel file
# - GET /files/{file_id}/preview - Preview file content and structure
# - PUT /files/{file_id}/metadata - Update file metadata (name, description)

from fastapi import APIRouter, Depends, HTTPException, status, UploadFile, File, Query
from fastapi.responses import StreamingResponse,RedirectResponse
from typing import List,Optional
import io
from datetime import datetime,UTC

from services.file_service import(
    validate_file_upload,upload_file,get_user_files,
    get_file_download_url,delete_file,get_file_by_id
)

from models.file import FileResponse,FileType,ProcessingStatus,FileUpload,TemplateCategory
from api.deps import(
    get_current_active_user,get_db,validate_file_upload as validate_file_upload_dep,
    get_pagination_params,require_credits
)
router=APIRouter(prefix="/files",tags=["Files"])

@router.post("/upload",response_model=FileResponse,status_code=status.HTTP_201_CREATED)
async def upload_excel_file(
    file:UploadFile=Depends(validate_file_upload_dep),
    current_user=Depends(require_credits(1)),
    db=Depends(get_db)
):
    """Upload Excel file with validation"""
    try:
        # Read file data
        file_data = await file.read()
        
        # Determine file type
        file_type = FileType.EXCEL
        if file.filename and file.filename.lower().endswith('.csv'):
            file_type = FileType.CSV
        
        # Create FileUpload object
        upload_request = FileUpload(
            filename=file.filename or "unknown.xlsx",
            file_size=len(file_data),
            file_type=file_type,
            template_category=TemplateCategory.BUSINESS  # Default category
        )
        
        # Call service with correct signature
        uploaded_file = await upload_file(
            user_id=str(current_user.id),
            file_data=file_data,
            upload_request=upload_request
        )
        
        return FileResponse.from_orm(uploaded_file)
    
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"File upload failed: {str(e)}"
        )
    

@router.get("",response_model=List[FileResponse])
async def list_user_files(
    pagination=Depends(get_pagination_params),
    status_filter:Optional[ProcessingStatus]=Query(None,description="Filter by processing status"),
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """List user's uploaded files with pagination"""
    try:
        files=await get_user_files(
            user_id=str(current_user.id),
            skip=pagination["skip"],
            limit=pagination["limit"],
            status_filter=status_filter
        )
        return [FileResponse.from_orm(f) for f in files]
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error retrieving files: {str(e)}"
        )
    

@router.get("/{file_id}",response_model=FileResponse)
async def get_file_details(
    file_id:str,
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """Get details of a specific uploaded file"""
    try:
        file=await get_file_by_id(file_id,str(current_user.id))
        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found"
            )
        return FileResponse.from_orm(file)
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error retrieving file details: {str(e)}"
        )
    

@router.delete("/{file_id}")
async def delete_user_file(
    file_id:str,
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """Delete an uploaded file"""
    try:
        file=await get_file_by_id(file_id,str(current_user.id))
        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found"
            )
        
        success=await delete_file(file_id,str(current_user.id))
        if not success:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="File deletion failed"
            )
        return {"message":"File deleted successfully","file_id":file_id,"file_name":file.name}
    

    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error deleting file: {str(e)}"
        )



@router.get("/{file_id}/download")
async def download_file(
    file_id:str,
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """Download the original Excel file"""
    try:
        file=await get_file_by_id(file_id,str(current_user.id))
        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found"
            )
        
        download_url=await get_file_download_url(file_id,str(current_user.id))
        if not download_url:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail="Could not generate download URL"
            )
        
        return RedirectResponse(url=download_url)
    
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error downloading file: {str(e)}"
)
    

@router.get("/{file_id}/preview")
async def preview_file_content(
    file_id:str,
    sheet_name:Optional[str]=Query(None,description="Specific sheet to preview"),
    rows_limit:int=Query(10,ge=1,le=100,description="Number of rows to preview"),
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """Preview file content and structure"""
    try:
        file=await get_file_by_id(file_id,str(current_user.id))
        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found"
            )
        
        preview_data={
            "file_id":file.id,
            "file_name":file.name,
            "file_type":file.file_type,
            "sheets":[
                {
                "sheet_name": "Sheet1",
                "columns":["Column A","Column B","Column C"],
                "rows":[
                    ["value1","value2","value3"],
                    ["value4","value5","value6"]
                ],
                "total_rows": 100,
                "preview_rows": rows_limit
                }
            ],
            "message":"Preview data is mocked for demonstration purposes"
        }
        return preview_data
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error previewing file: {str(e)}"
        )
    


@router.put("/{file_id}/metadata")
async def update_file_metadata(
    file_id:str,
    filename:Optional[str]=Query(None,min_length=1,max_length=255,description="New file name"),
    description:Optional[str]=Query(None,max_length=500,description="File description"),
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """Update file metadata (name, description)"""
    try:
        file=await get_file_by_id(file_id,str(current_user.id))
        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found"
            )
        
        update_data={}
        if filename:
            update_data["name"]=filename
        if description:
            update_data["description"]=description

        if not update_data:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="No valid fields provided for update"
            )

        return {
            "message":"File metadata updated successfully",
            "file_id":file_id,
            "updated_fields":update_data,
            "note":"Metadata update is mocked for demonstration purposes"
        }
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error updating file metadata: {str(e)}"
        )
    

@router.get("/{file_id}/stats")
async def get_file_statistics(
    file_id:str,
    current_user=Depends(get_current_active_user),
    db=Depends(get_db)
):
    """Get usage statistics for a specific file"""
    try:
        file=await get_file_by_id(file_id,str(current_user.id))
        if not file:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail="File not found"
            )
        
        stats={
            "file_id":file.id,
            "file_name":file.name,
            "file_size":file.size,
            "file_type":file.file_type,
            "upload_date":file.uploaded_at,
            "status":file.status,
            "processing stats":{
                "sheet_detected":1,
                "total_rows":0,
                "total_columns":0,
                "data_types_detected":[],
                "estimated_processing_time_seconds":"2-5 minutes"
            },
            "conversion_history":{
                "total_conversions":0,
                "successful_conversions":0,
                "last_conversion_date":None
            }
        }
        return stats
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Error retrieving file statistics: {str(e)}"
        )