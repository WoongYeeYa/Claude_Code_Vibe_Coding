@echo off
chcp 65001 >nul
cls
echo.
echo ============================================================
echo    Word to HWP Converter
echo ============================================================
echo.
echo   1. Install dependencies
echo   2. Start server
echo   3. Start server and open browser
echo   4. Stop server
echo   5. Exit
echo.
echo ============================================================
echo.

set /p choice=Select option (1-5):

if "%choice%"=="1" call install.bat
if "%choice%"=="2" call start_server.bat
if "%choice%"=="3" (
    start /B cmd /c start_server.bat
    timeout /t 3 /nobreak >nul
    call open_browser.bat
)
if "%choice%"=="4" call stop_server.bat
if "%choice%"=="5" exit

goto :EOF
