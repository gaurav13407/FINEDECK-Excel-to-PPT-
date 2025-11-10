@echo off
echo ============================================================
echo COMPLETE BACKEND RESTART PROCEDURE
echo ============================================================
echo.

echo Step 1: Killing ALL Python processes...
taskkill /F /IM python.exe /T 2>nul
if %ERRORLEVEL% EQU 0 (
    echo    SUCCESS: Python processes killed
) else (
    echo    INFO: No Python processes were running
)
timeout /t 2 /nobreak >nul

echo.
echo Step 2: Verifying Python is dead...
tasklist | findstr /I "python.exe" >nul
if %ERRORLEVEL% EQU 0 (
    echo    WARNING: Python still running! Trying again...
    taskkill /F /IM python.exe /T
    timeout /t 2 /nobreak >nul
) else (
    echo    SUCCESS: No Python processes found
)

echo.
echo Step 3: Deleting __pycache__ directories...
cd /d "c:\Users\gaura\OneDrive\Desktop\Big Projects\FinDeck(Excel to PPT Project)"
for /d /r . %%d in (__pycache__) do (
    if exist "%%d" (
        echo    Deleting: %%d
        rd /s /q "%%d" 2>nul
    )
)
echo    SUCCESS: Cache cleared

echo.
echo Step 4: Deleting .pyc files...
del /s /q *.pyc 2>nul
echo    SUCCESS: .pyc files deleted

echo.
echo Step 5: Starting backend...
cd src\backend
echo    Starting uvicorn on port 8000...
echo.
echo ============================================================
echo WATCH FOR THESE MESSAGES:
echo    - Using SimpleFinanceChartBuilder (NO AI)
echo    - Chart system: SimpleFinanceChartBuilder
echo ============================================================
echo.

uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
