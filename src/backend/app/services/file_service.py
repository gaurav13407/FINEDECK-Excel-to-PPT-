# File Management Service
# Handles file operations and storage
# - File upload validation and processing
# - Excel file parsing and sheet analysis
# - File storage to Azure Blob Storage
# - File metadata extraction and indexing
# - File sharing and permission management
# - Temporary file cleanup and storage optimization
# - File preview generation and thumbnail creation

from typing import Optional,Dict,List,Any,Tuple
from datetime import datetime,timedelta,UTC
from bson import ObjectId
import os
import tempfile
from pathlib import Path

from models.file import(
    FileUpload,FileRecord,FileResponse,ConversionJob,
    FileType,ProcessingStatus,TemplateCategory,
    validate_template_access,calculate_credits_needed,get_subscription_limits
)
from models.user import UserInDB,SubscriptionPlan
from core.database import get_collection
from storage.b2_storage import get_storage_client
from services.user_service import get_user_by_id,deduct_user_credits

MAX_FILE_SIZE=50*1024*1024  # 50 MB
ALLOWED_EXTENSIONS={".xlsx",".xls",".csv"}

async def validate_file_upload(user_id:str,filename:str,file_size:int,
                               template_category:TemplateCategory)->Tuple[bool,str]:
    """Validate if user upload a file"""
    try:
        user=await get_user_by_id(user_id)
        if not user:
            return False,"User not found."
        if not user.is_active:
            return False,"User account is inactive."
        
        file_ext=Path(filename).suffix.lower()
        if file_ext not in ALLOWED_EXTENSIONS:
            return False,f"File type {file_ext} is not allowed."
        
        if file_size<=0:
            return False,"File size must be greater than 0."
        
        if file_size>MAX_FILE_SIZE:
            return False,f"File size exceeds the maximum limit of {MAX_FILE_SIZE/(1024*1024)} MB."
        
        user_plan=user.subscription.plan if hasattr(user.subscription.plan,"value")else user.subscription.plan
        if not validate_template_access(template_category,user_plan):
            return False,f"Template category {template_category.value} not available for{user_plan} plan."
        
        limits=get_subscription_limits(user_plan)
        files_collection=get_collection("files")
        current_month_start=datetime.now(UTC).replace(day=1,hour=0,minute=0,second=0,microsecond=0)

        monthly_file_count=await files_collection.count_documents({
            "user_id":ObjectId(user_id),
            "created_at":{"$gte":current_month_start}
        })
        if monthly_file_count>=limits["monthly_file_limit"]:
            return False,f"Monthly file limit reached({limits['monthly_file_limit']}files)"
        return True,"File upload validated successfully."
    except Exception as e:
        return False,f"Validation error:{e}"
    



async def upload_file(user_id:str,file_data:bytes,upload_request:FileUpload)->Optional[FileRecord]:
    """Upload file to storage and create database record"""
    try:
        is_valid,error_msg=await validate_file_upload(
            user_id,upload_request.filename,upload_request.file_size,upload_request.template_category
        )
        if not is_valid:
            raise Exception(error_msg)
        storage_client=get_storage_client(use_b2=False)

        file_ext=Path(upload_request.filename).suffix
        unique_filename=f"{user_id}/{datetime.now(UTC).strftime('%Y/%m/%d')}/{ObjectId()}{file_ext}"


        storage_path=await storage_client.upload_file(file_data,unique_filename)

        user=await get_user_by_id(user_id)

        file_record=FileRecord(
            user_id=ObjectId(user_id),
            filename=unique_filename,
            original_filename=upload_request.filename,
            file_size=upload_request.file_size,
            file_type=upload_request.file_type,
            template_category=upload_request.template_category,
            template_name=upload_request.template_name,
            conversion_options=upload_request.conversion_options,
            status=ProcessingStatus.UPLOADED,
            storage_path=storage_path,
            subscription_plan=user.subscription.plan if hasattr(user.subscription.plan,"value")else user.subscription.plan,
            credits_used=calculate_credits_needed(upload_request.template_category),
            created_at=datetime.now(UTC),
            updated_at=datetime.now(UTC)

        )
        files_collection=get_collection("files")
        result=await files_collection.insert_one(file_record.dict(by_alias=True))
        file_record.id=str(result.inserted_id)
        return file_record
    
    except Exception as e:
        raise Exception(f"File upload error:{e}")
    

async def create_conversion_job(file_id:str,user_id:str,conversion_config:Dict[str,Any])->Optional[ConversionJob]:
    """Create a conversion job for Excel to PPT processing"""
    try:
        files_collection=get_collection("files")
        file_record=await files_collection.find_one({
            "_id":ObjectId(file_id),
            "user_id":ObjectId(user_id)
        })

        if not file_record:
            raise Exception("File not found for conversion.")
        
        if file_record["status"]!=ProcessingStatus.UPLOADED:
            raise Exception(f"File not ready for conversion.Current status:{file_record['status']}")
        
        credits_needed=file_record["credits_used"]
        success=await deduct_user_credits(user_id,credits_needed)
        if not success:
            raise ValueError("Failed to deduct credits for conversion.")
        
        job=ConversionJob(
            file_id=file_id,
            user_id=ObjectId(user_id),
            job_type="Excel_to_ppt",
            status=ProcessingStatus.QUEUED,
            conversion_config=conversion_config or {},
            template_config={
                "category":file_record["template_category"],
                "name":file_record["template_name"]
            },
            created_at=datetime.now(UTC),
        )

        jobs_collection=get_collection("conversion_jobs")
        result=await jobs_collection.insert_one(job.dict(by_alias=True))
        job.id=str(result.inserted_id)

        await files_collection.update_one(
            {"_id":ObjectId(file_id)},
            {
                "$set":{
                    "status":ProcessingStatus.QUEUED,
                    "updated_at":datetime.now(UTC)}
            }
        )
        return job
    except Exception as e:
        raise Exception(f"Conversion job creation error:{e}")
    

async def get_user_files(user_id: str, skip: int = 0, limit: int = 20, status_filter: Optional[ProcessingStatus] = None) -> List[FileResponse]:
    """Get user's files with pagination and filtering"""
    try:
        files_collection = get_collection("files")
        query = {"user_id": ObjectId(user_id)}
        if status_filter:
            query["status"] = status_filter
        
        cursor = files_collection.find(query).sort("created_at", -1).skip(skip).limit(limit)
        files=await cursor.to_list(length=limit)

        file_responses=[]
        for file_doc in files:
            file_response=FileResponse(
                id=str(file_doc["_id"]),
                filename=file_doc["original_filename"],
                file_size=file_doc["file_size"],
                file_type=file_doc["file_type"],
                template_category=file_doc["template_category"],
                template_name=file_doc.get("template_name"),
                status=file_doc["status"],
                download_url=file_doc.get("download_url"),
                uploaded_at=file_doc["created_at"],
                expires_at=file_doc["expires_at"],
                created_at=file_doc["created_at"],
                updated_at=file_doc["updated_at"],
                subscription_plan=file_doc["subscription_plan"],
                credits_used=file_doc["credits_used"],
                error_message=file_doc.get("error_message"),
                progress_percentage=file_doc.get("progress_percentage",0)
            )
            file_responses.append(file_response)
        return file_responses
    except Exception as e:
        raise Exception(f"Get user files error:{e}")
    


async def get_file_download_url(file_id:str,user_id:str,url_expiry_hours:int=24)->Optional[str]:
    """Generate Secure download URL for a file"""
    try:
        files_collection=get_collection("files")
        file_record=await files_collection.find_one({
            "_id":ObjectId(file_id),
            "user_id":ObjectId(user_id)
        })

        if not file_record:
            raise Exception("File not found.")
        
        if file_record["status"] not in [ProcessingStatus.COMPLETED,ProcessingStatus.UPLOADED]:
            raise ValueError(f"File not available for download.Current status:{file_record['status']}")
        
        if file_record["expires_at"]<datetime.now(UTC):
            raise ValueError("File has expired and is no longer available for download.")
        
        storage_client=get_storage_client(use_b2=True)

        storage_path=file_record.get("output_storage_path",file_record["storage_path"])

        download_url=await storage_client.generate_download_url(
            storage_path,
            expiry_hours=url_expiry_hours
        )

        await files_collection.update_one(
            {"_id":ObjectId(file_id)},
            {"$inc":{"download_count":1},
                "$set":{"downloaded_at":datetime.now(UTC),
                        "updated_at":datetime.now(UTC),
                        "download_url":download_url}
            }
        )
        return download_url
    except Exception as e:
        raise Exception(f"Generate download URL error:{str(e)}")
    


async def delete_file(file_id:str,user_id:str)->bool:
    """Delete a user's file from storage and database"""
    try:
        files_collection=get_collection("files")
        file_record=await files_collection.find_one({
            "_id":ObjectId(file_id),
            "user_id":ObjectId(user_id)
        })

        if not file_record:
            raise Exception("File not found.")
        
        storage_client=get_storage_client(use_b2=True)

        if file_record.get("storage_path"):
            await storage_client.delete_file(file_record["storage_path"])
        if file_record.get("output_storage_path"):
            await storage_client.delete_file(file_record["output_storage_path"])

        jobs_collection=get_collection("conversion_jobs")
        await jobs_collection.delete_many({"file_id":file_id})

        result=await files_collection.delete_one({"_id":ObjectId(file_id)})
        return result.deleted_count>0
    except Exception as e:
        raise Exception(f"Delete file error:{str(e)}")


async def get_file_by_id(file_id: str, user_id: str) -> Optional[FileRecord]:
    """Get a specific file by ID for a user"""
    try:
        files_collection = get_collection("files")
        file_record = await files_collection.find_one({
            "_id": ObjectId(file_id),
            "user_id": ObjectId(user_id)
        })
        
        if not file_record:
            return None
            
        return FileRecord(
            id=str(file_record["_id"]),
            filename=file_record["filename"],
            file_size=file_record["file_size"],
            file_type=file_record["file_type"],
            user_id=str(file_record["user_id"]),
            upload_date=file_record["upload_date"],
            status=file_record["status"],
            storage_path=file_record.get("storage_path"),
            download_url=file_record.get("download_url")
        )
    except Exception as e:
        raise Exception(f"Get file by ID error: {str(e)}")
    

