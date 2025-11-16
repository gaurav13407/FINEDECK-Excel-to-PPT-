@echo off
echo ========================================
echo Starting FinDeck Application
echo ========================================
echo.

REM Start Backend Server
echo [1/2] Starting Backend Server...
start "FinDeck Backend" cmd /k "cd /d "%~dp0src\backend" && python -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000"

REM Wait a bit for backend to start
timeout /t 3 /nobreak > nul

REM Start Frontend Server
echo [2/2] Starting Frontend Server...
start "FinDeck Frontend" cmd /k "cd /d "%~dp0src\ui" && python -m http.server 8001"

REM Wait a bit for frontend to start
timeout /t 2 /nobreak > nul

echo.
echo ========================================
echo FinDeck Started Successfully!
echo ========================================
echo.
echo Backend:  http://localhost:8000
echo Frontend: http://localhost:8001
echo.
echo Opening browser...
start http://localhost:8001/mainpage.html
echo.
echo Press any key to stop all servers...
pause > nul

REM Kill both servers when user presses a key
taskkill /FI "WindowTitle eq FinDeck Backend*" /T /F
taskkill /FI "WindowTitle eq FinDeck Frontend*" /T /F
