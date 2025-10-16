@echo off
chcp 65001 >nul
cls
echo.
echo ============================================================
echo    Word to HWP Converter - Installation
echo ============================================================
echo.

cd /d "%~dp0"

echo Checking Python installation...
python --version
if errorlevel 1 (
    echo.
    echo [ERROR] Python is not installed.
    echo Please install Python 3.8 or higher.
    echo.
    pause
    exit /b 1
)
echo.

echo Upgrading pip...
python -m pip install --upgrade pip
echo.

echo Installing dependencies...
cd backend
pip install -r requirements.txt
echo.

echo ============================================================
echo Installation completed successfully!
echo.
echo Run start_server.bat to start the server.
echo ============================================================
echo.
pause
