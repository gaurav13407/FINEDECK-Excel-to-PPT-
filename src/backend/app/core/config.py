# Application Configuration Settings
# This file manages all environment variables and app settings
# - Database connection URLs (MongoDB, Redis)
# - JWT secret keys and token expiration times
# - File upload limits and allowed file types
# - Azure storage account configuration
# - API rate limiting settings
# - Environment-specific settings (dev, staging, prod)
# - CORS origins for frontend domains
from pydantic_settings import BaseSettings
from typing import Optional
import os

class Settings(BaseSettings):
    # Database URL -required,no default
    database_url:str
    database_name:str="findeck_db"
    # JWT Secret- required for production, has development default
    jwt_secret:str="dev-secret-change-in-production"
    jwt_algorithm:str="HS256"
    jwt_expiration_minutes:int =1440
    # Deubg mode-Optional.defaults to false
    debug:bool =False
    #Azure Storage Settings:
    azure_storage_connection_string:str
    azure_container_name:str="user-files"

    # Whats the max file size?(in MB)
    max_file_size_mb:int=50
    # What file typed allowed?
    allowed_file_extensions:list=[".xlsx",".xls"]

    ## App Information
    app_name:str="FinDeck API"
    api_version:str="v1"
    environment:str="development"

    #CORS Setting for frontend
    cors_origins:list=["http://localhost:3000", "http://127.0.0.1:5500","https://www.findeck.live"]
    # Backblaze B2 Storage Settings
    use_b2_storage: bool = False
    b2_application_key_id: str
    b2_application_key: str
    b2_bucket_name: str = "findeck-files"
    b2_endpoint: str = "s3.us-east-005.backblazeb2.com"

    # File Cleanup Settings
    cleanup_enabled: bool = True
    file_retention_hours: int = 24
    cleanup_interval_hours: int = 6





    class Config:
        env_file=".env"

# Create instance that other files can import (OUTSIDE the class)
settings = Settings()

    