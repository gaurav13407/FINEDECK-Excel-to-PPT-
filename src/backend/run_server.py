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
    
    print("Starting FinDeck FastAPI Server...")
    print(f"Working directory: {os.getcwd()}")
    print(f"Python path includes: {app_dir}")
    print("Server will be available at: http://localhost:8000")
    print("API documentation at: http://localhost:8000/docs")
    
    # Start server with import string for proper reload functionality
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=8000,
        reload=True,
        reload_dirs=[str(app_dir)]
    )