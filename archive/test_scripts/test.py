# Create this file: test_connection.py (in your root directory)
from src.backend.app.core.config import settings
print("✅ Config loaded successfully!")
print(f"Database URL: {settings.database_url}")
print(f"Database Name: {settings.database_name}")
print(f"JWT Secret: {settings.jwt_secret[:10]}...")  # Only show first 10 chars