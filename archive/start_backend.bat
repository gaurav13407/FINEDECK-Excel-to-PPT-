@echo off
echo ========================================
echo   Starting FinDeck Backend Server
echo ========================================
echo.

cd src\backend\app

echo Checking Python environment...
python --version
echo.

echo Starting FastAPI server on http://localhost:8000...
echo.
echo Press Ctrl+C to stop the server
echo.

python main.py

pause
