@echo off
chcp 65001 >nul
cls
echo.
echo ============================================================
echo    Word to HWP Converter - Server Starting
echo ============================================================
echo.

cd /d "%~dp0"

echo Checking Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed.
    echo Please install Python 3.8 or higher.
    echo.
    pause
    exit /b 1
)

echo Starting server...
echo.
echo ============================================================
echo Server URL: http://localhost:8000
echo API Docs:   http://localhost:8000/docs
echo.
echo Press Ctrl+C to stop the server
echo ============================================================
echo.

cd backend
python main.py

pause
