@echo off
echo ========================================
echo FinDeck - Install Dependencies
echo ========================================
echo.

echo [1/3] Activating virtual environment...
call .venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Virtual environment not found!
    echo Please create it first: python -m venv .venv
    pause
    exit /b 1
)
echo ✓ Virtual environment activated
echo.

echo [2/3] Installing SlowAPI for rate limiting...
pip install slowapi==0.1.9
if errorlevel 1 (
    echo ERROR: Failed to install slowapi
    pause
    exit /b 1
)
echo ✓ SlowAPI installed successfully
echo.

echo [3/3] Verifying installation...
python -c "import slowapi; print('✓ SlowAPI version:', slowapi.__version__)"
if errorlevel 1 (
    echo ERROR: SlowAPI import failed
    pause
    exit /b 1
)
echo.

echo ========================================
echo Installation Complete! ✅
echo ========================================
echo.
echo Next steps:
echo 1. Review .env file for security changes
echo 2. Run database migration (optional):
echo    python scripts\migrate_subscriptions.py --dry-run
echo 3. Start backend:
echo    cd src\backend\app
echo    uvicorn main:app --reload --host 0.0.0.0 --port 8000
echo 4. Visit API docs:
echo    http://localhost:8000/api/docs
echo.
pause
