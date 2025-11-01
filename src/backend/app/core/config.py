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
from pydantic import field_validator
from typing import Optional, Union
import os
import json
from pathlib import Path

# Get the project root directory (where .env file should be located)
current_file = Path(__file__)
project_root = current_file.parent.parent.parent.parent.parent  # Go up 5 levels from src/backend/app/core/config.py
env_file_path = project_root / ".env"

class Settings(BaseSettings):
    # Database URL - optional for development
    database_url: Optional[str] = "mongodb://localhost:27017"
    database_name: str = "findeck_db"
    # JWT Secret- required for production, has development default
    jwt_secret: str = "dev-secret-change-in-production"
    jwt_algorithm: str = "HS256"
    jwt_expiration_minutes: int = 1440
    # Debug mode-Optional.defaults to false
    debug: bool = False
    # Azure Storage Settings - optional
    azure_storage_connection_string: Optional[str] = None
    azure_container_name: str = "user-files"

    # Whats the max file size?(in MB)
    max_file_size_mb:int=50
    # What file typed allowed?
    allowed_file_extensions:list=[".xlsx",".xls"]

    ## App Information
    app_name:str="FinDeck API"
    api_version:str="v1"
    environment:str="development"

    #CORS Setting for frontend
    cors_origins: Union[list, str] = ["http://localhost:3000", "http://127.0.0.1:5500", "http://localhost:5500", "null", "https://www.findeck.live"]
    
    @field_validator('cors_origins', mode='before')
    @classmethod
    def parse_cors_origins(cls, v):
        """Parse CORS_ORIGINS from JSON string in .env file"""
        if isinstance(v, str):
            try:
                # Try to parse as JSON array
                parsed = json.loads(v)
                if isinstance(parsed, list):
                    # Add 'null' origin for local file:// access if not present
                    if "null" not in parsed:
                        parsed.append("null")
                    return parsed
            except json.JSONDecodeError:
                # If not valid JSON, split by comma
                origins = [origin.strip().strip('"').strip("'") for origin in v.split(',')]
                if "null" not in origins:
                    origins.append("null")
                return origins
        elif isinstance(v, list):
            # Already a list, just ensure 'null' is included
            if "null" not in v:
                v.append("null")
            return v
        return v
    
    # Backblaze B2 Storage Settings
    use_b2_storage: bool = False
    b2_application_key_id: Optional[str] = None
    b2_application_key: Optional[str] = None
    b2_bucket_name: str = "findeck-files"
    b2_endpoint: str = "s3.us-east-005.backblazeb2.com"

    # File Cleanup Settings
    cleanup_enabled: bool = True
    file_retention_hours: int = 24
    cleanup_interval_hours: int = 6

    # Transactional email providers
    brevo_api_key: Optional[str] = None
    sendgrid_api_key: Optional[str] = None
    email_from: str = "no-reply@example.com"
    frontend_url: str = "http://localhost:3000"

    model_config = {"env_file": str(env_file_path), "extra": "ignore"}

    # OAuth provider credentials (uppercase names are read by oauth.py via getattr)
    GOOGLE_CLIENT_ID: Optional[str] = None
    GOOGLE_CLIENT_SECRET: Optional[str] = None
    GOOGLE_CLIENT_ID_LOCAL: Optional[str] = None
    GOOGLE_CLIENT_SECRET_LOCAL: Optional[str] = None
    GOOGLE_CLIENT_ID_PROD: Optional[str] = None
    GOOGLE_CLIENT_SECRET_PROD: Optional[str] = None

    GITHUB_CLIENT_ID: Optional[str] = None
    GITHUB_CLIENT_SECRET: Optional[str] = None
    GITHUB_CLIENT_ID_LOCAL: Optional[str] = None
    GITHUB_CLIENT_SECRET_LOCAL: Optional[str] = None
    GITHUB_CLIENT_ID_PROD: Optional[str] = None
    GITHUB_CLIENT_SECRET_PROD: Optional[str] = None

    X_CLIENT_ID: Optional[str] = None
    X_CLIENT_SECRET: Optional[str] = None
    X_CLIENT_ID_LOCAL: Optional[str] = None
    X_CLIENT_SECRET_LOCAL: Optional[str] = None
    X_CLIENT_ID_PROD: Optional[str] = None
    X_CLIENT_SECRET_PROD: Optional[str] = None

# Create instance that other files can import (OUTSIDE the class)
settings = Settings()

# Debug: Print the loaded settings
print(f"🔧 Config Debug:")
print(f"   .env file path: {env_file_path}")
print(f"   .env file exists: {env_file_path.exists()}")
print(f"   DATABASE_URL: {settings.database_url}")
print(f"   DATABASE_NAME: {settings.database_name}")
print(f"   CORS_ORIGINS: {settings.cors_origins}")
print(f"   DEBUG: {settings.debug}")
print(f"   ENVIRONMENT: {settings.environment}")
try:
    google_loaded = bool(getattr(settings, "GOOGLE_CLIENT_ID", None))
    github_loaded = bool(getattr(settings, "GITHUB_CLIENT_ID", None))
    x_loaded = bool(getattr(settings, "X_CLIENT_ID", None))
    print(f"   GOOGLE_CLIENT_ID present: {google_loaded}")
    print(f"   GITHUB_CLIENT_ID present: {github_loaded}")
    print(f"   X_CLIENT_ID present: {x_loaded}")
except Exception:
    pass

    