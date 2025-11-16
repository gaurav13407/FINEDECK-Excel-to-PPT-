@echo off
echo ========================================
echo Testing Backend Server
echo ========================================
echo.

cd /d "%~dp0src\backend"

echo Starting backend server with detailed logging...
echo.
echo Press Ctrl+C to stop the server
echo.

python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000 --log-level debug

pause
