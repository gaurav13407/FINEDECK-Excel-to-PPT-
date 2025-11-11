#!/usr/bin/env python3
"""
Server startup script for FinDeck FastAPI application.
This script properly sets up the Python path and starts the server.
"""

import os
import sys
from pathlib import Path

# Add the app directory to Python path
current_dir = Path(__file__).parent
app_dir = current_dir / "app"
sys.path.insert(0, str(app_dir))

# Change working directory to app directory
os.chdir(app_dir)

if __name__ == "__main__":
    import uvicorn
    
    # Get port from environment (for Render/production) or use default
    port = int(os.getenv("PORT", 8000))
    
    # Check if running in production
    is_production = os.getenv("ENVIRONMENT", "development") == "production"
    
    print("Starting FinDeck FastAPI Server...")
    print(f"Environment: {'Production' if is_production else 'Development'}")
    print(f"Working directory: {os.getcwd()}")
    print(f"Python path includes: {app_dir}")
    print(f"Server will be available on port: {port}")
    if not is_production:
        print(f"Server will be available at: http://localhost:{port}")
        print(f"API documentation at: http://localhost:{port}/docs")
    
    # Start server with import string for proper reload functionality
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=port,
        reload=not is_production,  # Only reload in development
        reload_dirs=[str(app_dir)] if not is_production else None
    )