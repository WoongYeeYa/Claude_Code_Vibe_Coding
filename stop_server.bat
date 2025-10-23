@echo off
echo.
echo ========================================
echo  DOCX to HWP - Server Stop
echo ========================================
echo.

echo Checking for Java processes...
tasklist | findstr /I "java"

echo.
echo Stopping all Java processes...
taskkill /F /IM java.exe /T 2>nul
taskkill /F /IM javaw.exe /T 2>nul

if errorlevel 1 (
    echo No Java process found
) else (
    echo All Java processes stopped
)

echo.
echo Done!
pause
