# File MongoDB Model
# Handles all file-related data structures
# - File metadata (name, size, type, upload date)
# - Excel file structure and sheet information
# - PowerPoint template associations
# - Conversion status and processing history
# - File storage paths and cloud URLs
# - Download links and expiration times
# - File sharing permissions and access control

from datetime import datetime,timedelta
from typing import Optional,List,Dict,Any
from enum import Enum
from pydantic import BaseModel,Field,validator
from bson import ObjectId
from .user import TemplateCategory, PyObjectId, SubscriptionPlan


class FileType(str,Enum):
    """Supported excel file types"""
    EXCEL="excel"
    CSV="csv"
    XLSX="xlsx"
    XLS="xls"

class ProcessingStatus(str,Enum):
    """File processing status"""
    UPLOADED="uploaded"
    QUEUED="queued"
    PROCESSING="processing"
    COMPLETED="completed"
    FAILED="failed"
    EXPIRED="expired"


class StorageProvider(str,Enum):
    """"Storage backend options"""
    LOCAL="local"
    AZURE_BLOB="azure_blob"


class FileUpload(BaseModel):
    """Model for file upload request"""
    filename:str=Field(...,min_length=1,max_length=255)
    file_size:int=Field(...,gt=0,le=50*1024*1024)
    file_type:FileType
    template_category:TemplateCategory
    template_name:Optional[str]=Field(None,max_length=100)
    conversion_options:Optional[Dict[str,Any]]=Field(default_factory=dict)

    @validator('filename')
    def validate_filename(cls,v):
        """Vaildate File Name"""
        allowed_extensions={".xlsx",".xls",".csv"}
        if not any(v.lower().endswith(ext) for ext in allowed_extensions):
            raise ValueError(f"Filename must end with one of {allowed_extensions}")
        return v
    

class FileRecord(BaseModel):
    """Complete file record in DB storage"""
    id:Optional[str]=Field(default=None,alias="_id")
    user_id:PyObjectId=Field(...,description="User who uploaded the file")
    filename:str
    original_filename:str
    file_size:int
    file_type:FileType
    template_category:TemplateCategory
    template_name:Optional[str]=None
    conversion_options:Dict[str,Any]=Field(default_factory=dict)


    # Processing Tracking
    status:ProcessingStatus=ProcessingStatus.UPLOADED
    processing_started_at:Optional[datetime]=None
    processing_completed_at:Optional[datetime]=None
    error_message:Optional[str]=None
    progress_percentage:int=Field(default=0,ge=0,le=100)

    # Storage Info
    storage_provider:StorageProvider=StorageProvider.LOCAL
    storage_path:Optional[str]=None
    azure_blob_url:Optional[str]=None
    download_url:Optional[str]=None

    # file lifecycle
    uploaded_at:datetime=Field(default_factory=datetime.utcnow)
    expires_at:datetime=Field(default_factory=lambda:datetime.utcnow()+timedelta(hours=24))
    downloaded_at:Optional[datetime]=None
    download_count:int=Field(default=0,ge=0)

    ## Output file info
    output_filename:Optional[str]=None
    output_file_size:Optional[int]=None
    output_storage_path:Optional[str]=None

    # Subscription tracking
    subscription_plan:str=Field(...,description="User's subscription tier when file was processsed")
    credits_used:int=Field(default=1,ge=0,description="credits consumed for this conversion")

    # Excel file analysis
    excel_metadata:Optional[Dict[str,Any]]=Field(default_factory=dict)
    sheets_info:Optional[List[Dict[str,Any]]]=Field(default_factory=list)
    data_summary:Optional[Dict[str,Any]]=Field(default_factory=dict)

    # Metadata
    processing_metadata:Dict[str,Any]=Field(default_factory=dict)
    created_at:datetime=Field(default_factory=datetime.utcnow)
    updated_at:datetime=Field(default_factory=datetime.utcnow)

    class Config:
        populate_by_name=True
        json_encoders={
            ObjectId:str,
            datetime:lambda v:v.isoformat()
        }


    
class FileResponse(BaseModel):
    """Response model for file operations"""
    id:str
    filename:str
    file_size:int
    file_type:FileType
    template_category:TemplateCategory
    template_name:Optional[str]=None
    status:ProcessingStatus
    download_url:Optional[str]=None
    uploaded_at:datetime
    expires_at:datetime
    created_at:datetime
    updated_at:datetime
    subscription_plan:str
    credits_used:int
    error_message:Optional[str]=None
    progress_percentage:int


class ConversionJob(BaseModel):
    """Background job tracking model"""
    id:Optional[str]=Field(default=None,alias="_id")
    file_id:str
    user_id:PyObjectId
    job_type:str="excel_to_ppt"
    priority:int=Field(default=1,ge=1,le=5)

    # Job status
    status:ProcessingStatus=ProcessingStatus.QUEUED
    created_at:datetime=Field(default_factory=datetime.utcnow)
    started_at:Optional[datetime]=None
    completed_at:Optional[datetime]=None

    # Processing details
    worker_id:Optional[str]=None
    retry_count:int=Field(default=0,ge=0)
    max_retries:int=Field(default=3,ge=0)
    error_details:Optional[Dict[str,Any]]=None

    # Job Configuration (for my conversion logic)
    conversion_config:Dict[str,Any]=Field(default_factory=dict)
    template_config:Dict[str,Any]=Field(default_factory=dict)


    class Config:
        populate_by_name=True
        json_encoders={
            ObjectId:str,
            datetime:lambda v:v.isoformat()
        }


class FileUsageStats(BaseModel):
    """Usage stats for user files"""
    total_files_uploaded:int
    total_file_processing:int
    file_this_month:int
    credits_used_this_month:int
    monthly_file_limit:int
    monthly_credits_limit:int
    storage_used_mb:float
    storage_limit_mb:float
    sucess_rate_percentage:float


# utility functions for file management
def validate_template_access(template_category:TemplateCategory,user_subscription:str)->bool:
    """Validate if user can access requested template category"""
    access_map={
        "basic":[TemplateCategory.BASIC],
        "pro":[TemplateCategory.BASIC,TemplateCategory.PROFESSIONAL],
        "enterprise":[TemplateCategory.BASIC,TemplateCategory.PROFESSIONAL,TemplateCategory.PREMIUM,TemplateCategory.CUSTOM]
    }
    return template_category in access_map.get(user_subscription,[])

def calculate_credits_needed(template_category:TemplateCategory)->int:
    """Calculate credits based on template category"""
    credit_map={
        TemplateCategory.BASIC:1,
        TemplateCategory.PROFESSIONAL:2,
        TemplateCategory.PREMIUM:3,
        TemplateCategory.CUSTOM:5
    }
    return credit_map.get(template_category,1)

def get_subscription_limits(subscription_plan:str)->Dict[str,Any]:
    """Get limit for user's subscription plan"""
    limits={
        "basic":{"monthly_file_limit":7,"monthly_credits_limit":10,"storage_limit_mb":100},
        "pro":{"monthly_file_limit":15,"monthly_credits_limit":100,"storage_limit_mb":500},
        "enterprise":{"monthly_file_limit":1000,"monthly_credits_limit":5000,"storage_limit_mb":2000}
    }
    return limits.get(subscription_plan,limits["basic"])