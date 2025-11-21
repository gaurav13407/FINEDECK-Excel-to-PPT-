@echo off
echo ================================================================================
echo Starting FinDeck Server with Redis Session Middleware
echo ================================================================================
echo.
cd /d "%~dp0"
cd src\backend
python run_server.py
pause
