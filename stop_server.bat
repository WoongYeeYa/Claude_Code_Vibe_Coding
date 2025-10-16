@echo off
chcp 65001 >nul
cls
echo.
echo ============================================================
echo    Word to HWP Converter - Stopping Server
echo ============================================================
echo.

echo Checking for running Python processes...
tasklist | findstr python.exe >nul
if errorlevel 1 (
    echo No Python server is running.
) else (
    echo Stopping Python server...
    taskkill /F /IM python.exe /T >nul 2>&1
    echo Server stopped successfully.
)

echo.
pause
