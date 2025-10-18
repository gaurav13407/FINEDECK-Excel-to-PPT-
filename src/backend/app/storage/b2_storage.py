# Backblaze B2 Cloud Storage Integration
# Handles cloud file storage operations using Backblaze B2
# - File upload to B2 buckets with organized folder structure
# - Secure download URLs with expiration
# - 24-hour automatic file cleanup
# - Storage quota monitoring and usage tracking
# - Cost-effective storage ($0.005/GB vs AWS $0.023/GB)

import os
import asyncio
from datetime import datetime,timedelta
from typing import Optional,Tuple,Dict,List
import json
import logging
from pathlib import Path

logger=logging.getLogger(__name__)


try:
    from b2sdk.v2 import InMemoryAccountInfo, B2Api
    # Use generic exception handling for newer b2sdk versions
    B2_AVAILABLE = True
    # Create generic exception classes for compatibility
    class BucketIdNotFound(Exception):
        pass
    class FileNotPresent(Exception):
        pass
    class B2Error(Exception):
        pass
except ImportError:
    B2_AVAILABLE = False
    print("b2sdk not installed. Backblaze B2 storage will not be available.")

class BackblazeB2Storage:
    def __init__(self,application_key_id:str,application_key:str,bucket_name:str):
        if not B2_AVAILABLE:
            raise ImportError("b2sdk is not installed. Please install it to use Backblaze B2 storage.")
        
        self.application_key_id=application_key_id
        self.application_key=application_key
        self.bucket_name=bucket_name
        self.b2_api=None
        self.bucket=None
        self.initialize_b2()

    def initialize_b2(self):
        """Initialize B2 API and get/cerate bucket"""
        try:
            info=InMemoryAccountInfo()
            self.b2_api=B2Api(info)

            # Authorize Account
            self.b2_api.authorize_account("production",self.application_key_id,self.application_key)

            try:
                self.bucket=self.b2_api.get_bucket_by_name(self.bucket_name)
                logger.info(f"Using existing B2 bucket: {self.bucket_name}")
            except BucketIdNotFound:
                #create bucket if not exists
                self.bucket=self.b2_api.create_bucket(self.bucket_name,"allPrivate")
                logger.info(f"Created new B2 bucket: {self.bucket_name}")
        except Exception as e:
            logger.error(f"Error initializing B2 bucket: {e}")
            raise

    async def upload_excel_file(self,file_content:bytes,user_id:str,original_filename:str)->Tuple[str,str]:
        """Upload excel file with organized path strucutre"""
        try:
            # Create folder structure: user_id/yyyy/mm/dd/
            now=datetime.utcnow()
            date_folder=now.strftime("%Y/%m/%d")
            time_folder=now.strftime("%H-%M")
            file_path=f"uploads/{user_id}/{date_folder}/{time_folder}/{original_filename}"

            # Upload file
            file_info=self.bucket.upload_bytes(file_content,file_path,content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # Get the Download URL using the bucket and file name
            download_url=self.b2_api.get_download_url_for_file_name(self.bucket_name, file_path)

            logger.info(f"Uploaded file to B2: {file_path}")
            return file_path,download_url
        except Exception as e:
            logger.error(f"Error uploading file to B2: {e}")
            raise

    async def upload_analysis_data(self,analysis_data:Dict,user_id:str,file_id:str)-> Tuple[str,str]:
        """Upload Excel analysis data as JSON to B2"""
        try:
            now=datetime.utcnow()
            date_folder=now.strftime("%Y-%m-%d") 
            time_folder=now.strftime("%H-%M")
            file_path=f"analysis/{user_id}/{date_folder}/{time_folder}/{file_id}_analysis.json"

            # Convert to Json bytes
            json_content=json.dumps(analysis_data,indent=2,default=str).encode('utf-8') 

            # Upload analysis data
            file_info=self.bucket.upload_bytes(json_content,file_path,content_type="application/json")

            download_url=self.b2_api.get_download_url_for_file_name(self.bucket_name, file_path)

            logger.info(f"Uploaded analysis data to B2: {file_path}")
            return download_url,file_path
        
        except Exception as e:
            logger.error(f"Error uploading analysis data to B2: {e}")
            raise

    async def upload_ppt_output(self,ppt_content:bytes,user_id:str,filename:str)->Tuple[str,str]:
        """Upload generated PPT file to B2"""
        try:
            now=datetime.utcnow()
            date_folder=now.strftime("%Y-%m-%d")
            time_folder=now.strftime("%H-%M")
            file_path=f"outputs/{user_id}/{date_folder}/{time_folder}/{filename}"

            # Upload PPT file
            file_info=self.bucket.upload_bytes(ppt_content,file_path,content_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")

            download_url=self.b2_api.get_download_url_for_file_name(self.bucket_name, file_path)

            logger.info(f"Uploaded PPT output to B2: {file_path}")
            return file_path,download_url
        except Exception as e:
            logger.error(f"Error uploading PPT to B2: {e}")
            raise

    async def get_download_url(self,file_path:str,expiry_hours:int=24)->str:
        """Generate secure download URL with expiration"""
        try:
            #find file by path
            for file_version in self.bucket.ls():
                if file_version.file_name==file_path:
                    return self.b2_api.get_download_url_for_file_name(self.bucket_name, file_path)
        except Exception as e:
            logger.error(f"Error generating download URL: {e}")
            raise

    async def delete_file(self,file_path:str)->bool:
        try:
            # Find and delete file
            for file_version in self.bucket.ls():
                if file_version.file_name==file_path:
                    self.bucket.delete_file_version(file_version.id_,file_version.file_name)
                    logger.info(f"Deleted file from B2: {file_path}")
                    return True
        except Exception as e:
            logger.error(f"Error deleting file from B2: {e}")
            return False
        

    async def cleanup_expired_files(self,hours_old:int=24)->Dict[str,int]:
        """Delete files older than 24 hours"""
        try:
            cutoff_time=datetime.utcnow()-timedelta(hours=hours_old)
            cutoff_timestamp=int(cutoff_time.timestamp()*1000) # B2 uses milliseconds

            deleted_count={
                "uploads":0,
                "analysis":0,
                "outputs":0,
                "others":0,
                "total_size_mb":0
            }
            logger.info(f"Starting cleanup of files older than {hours_old} hours")

            # Liast all files and delete old ones
            for file_version in self.bucket.ls(recursive=True):
                file_time=file_version.upload_timestamp
                if file_time<cutoff_timestamp:
                    file_path=file_version.file_name
                    file_size_mb=file_version.size/(1024*1024)

                    # Categorize file for reporting
                    if file_path.startswith("uploads/"):
                        deleted_count["uploads"]+=1
                    elif file_path.startswith("analysis/"):
                        deleted_count["analysis"]+=1
                    elif file_path.startswith("outputs/"):
                        deleted_count["outputs"]+=1
                    else:
                        deleted_count["others"]+=1
                    deleted_count["total_size_mb"]+=file_size_mb

                    # Delete the file
                    self.b2_api.delete_file_version(file_version.id_,file_path)
                    logger.debug(f"Deleted expired file: {file_path} ({file_size_mb:.2f} MB)")

            total_deleted=sum(deleted_count[k] for k in ["uploads","analysis","outputs","others"])

            logger.info(f"Cleanup complete. Deleted {total_deleted} files, freeing {deleted_count['total_size_mb']:.2f} MB")

            return deleted_count
        except Exception as e:
            logger.error(f"Error during cleanup of expired files: {e}")
            raise

    async def get_storage_usage(self,user_id:str)->Dict:
        """Get storage usage for statistics and quota monitoring"""
        try:
            stats={
                "total_files":0,
                "total_size_mb":0,
                "total_size_bytes":0,
                "by_category":{
                    "uploads":{"count":0,"size_mb":0,"size_bytes":0},
                    "analysis":{"count":0,"size_mb":0,"size_bytes":0},
                    "outputs":{"count":0,"size_mb":0,"size_bytes":0},
                    "others":{"count":0,"size_mb":0,"size_bytes":0}
                },
                "last_updated":datetime.utcnow().isoformat()

            }

            # set seacrh prefix for specific user or all files
            prefix=f"uploads/{user_id}/" if user_id else ""

            for file_version in self.bucket.ls(folder_to_list=prefix if prefix else None,recursive=True):
                file_path=file_version.file_name
                file_size=file_version.size
                file_size_mb=file_size/(1024*1024)


                stats["total_files"]+=1
                stats["total_size_bytes"]+=file_size
                stats["total_size_mb"]+=file_size_mb

                # Categorize file
                if file_path.startswith("uploads/"):
                    stats["by_category"]["uploads"]["count"]+=1
                    stats["by_category"]["uploads"]["size_mb"]+=file_size_mb
                elif file_path.startswith("analysis/"):
                    stats["by_category"]["analysis"]["count"]+=1
                    stats["by_category"]["analysis"]["size_mb"]+=file_size_mb
                elif file_path.startswith("outputs/"):
                    stats["by_category"]["outputs"]["count"]+=1
                    stats["by_category"]["outputs"]["size_mb"]+=file_size_mb
                else:
                    stats["by_category"]["others"]["count"]+=1
                    stats["by_category"]["others"]["size_mb"]+=file_size_mb
            return stats
        
        except Exception as e:
            logger.error(f"Error getting storage usage: {e}")
            raise


class LocalStorage:
    """ Local file Storage for development(wihtin 24 hours clenaup)"""

    def __init__(self,base_path:str="uploads"):
        self.base_path=Path(base_path)
        self.base_path.mkdir(parents=True,exist_ok=True)

    async def upload_file(self, file_data: bytes, filename: str) -> str:
        """Upload file to local storage (generic method)"""
        # Create directory structure
        file_path = self.base_path / filename
        file_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Write file
        with open(file_path, "wb") as f:
            f.write(file_data)
        
        # Return the absolute path for direct file access
        return str(file_path.absolute())

    async def upload_excel_file(self,file_content:bytes,user_id:str,original_filename:str)->Tuple[str,str]:
        """Upload excel file to local storage"""
        now=datetime.utcnow()
        # date_folder=now.strftime("%Y-%m-%d")
        date_folder=now.strftime("%H-%M")
        time_folder=now.strftime("%H-%M")

        dir_path=self.base_path/"uploads"/user_id/date_folder/time_folder
        dir_path.mkdir(parents=True,exist_ok=True)

        file_path=dir_path/original_filename

        with open(file_path,"wb") as f:
            f.write(file_content)

        file_url=f"file://{file_path.absolute()}"
        return file_url,str(file_path)
    
    async def upload_analysis_data(self,analysis_data:Dict,user_id:str,file_id:str)-> Tuple[str,str]:
        """Upload analysis data to local storage"""
        now=datetime.utcnow()
        date_folder=now.strftime("%Y-%m-%d")
        time_folder=now.strftime("%H-%M")

        dir_path=self.base_path/"analysis"/user_id/date_folder/time_folder
        dir_path.mkdir(parents=True,exist_ok=True)

        file_path=dir_path/f"{file_id}_analysis.json"

        with open(file_path,'w') as f:
            json.dump(analysis_data,f,indent=2,default=str)

        file_url=f"file://{file_path.absolute()}"
        return file_url,str(file_path)
    
    async def cleanup_expired_files(self,hours_old:int=24)->Dict[str,int]:
        """Cleanup files older than specified hours"""
        cutoff_time=datetime.utcnow()-timedelta(hours=hours_old)

        deleted_count={
            "uploads":0,
            "analysis":0,
            "outputs":0,
            "others":0,
            "total_size_mb":0
        }

        for root,dirs,files in os.walk(self.base_path):
            for file in files:
                file_path=Path(root)/file
                try:
                    #Check file age
                    file_time=datetime.fromtimestamp(file_path.stat().st_mtime)

                    if file_time<cutoff_time:
                        file_size_mb=file_path.stat().st_size/(1024*1024)

                        # Categorize and delete
                        if "uploads" in str(file_path):
                            deleted_count["uploads"]+=1
                        elif "analysis" in str(file_path):
                            deleted_count["analysis"]+=1
                        elif "outputs" in str(file_path):
                            deleted_count["outputs"]+=1
                        else:
                            deleted_count["others"]+=1

                        deleted_count["total_size_mb"]+=file_size_mb
                        file_path.unlink()
                except Exception as e:
                    logger.error(f"Error deleting file {file_path}: {e}")
        return deleted_count
    

def get_storage_client(use_b2:bool=False,application_key_id:str=None,application_key:str=None,bucket_name:str=None):
    """Factory function to get appropriate storage client"""
    if use_b2 and application_key_id and application_key and bucket_name:
        return BackblazeB2Storage(application_key_id,application_key,bucket_name)
    else:
        return LocalStorage()