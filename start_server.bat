@echo off
echo.
echo ========================================
echo  DOCX to HWP - Server Start
echo ========================================
echo.

echo Current directory: %CD%
echo.

if not exist "backend-java\target\docx-to-hwp-converter-1.0.0.jar" (
    echo ========================================
    echo ERROR: JAR file not found!
    echo ========================================
    echo.
    echo Please run build.bat or install.bat first
    echo.
    echo Expected location:
    echo backend-java\target\docx-to-hwp-converter-1.0.0.jar
    echo.
    pause
    exit /b 1
)

echo Starting server...
echo.
echo ========================================
echo  Server: http://localhost:8000
echo.
echo  Press Ctrl+C to stop
echo ========================================
echo.

cd backend-java
"C:\Program Files\Java\jdk-17\bin\java.exe" -jar target\docx-to-hwp-converter-1.0.0.jar
cd ..

echo.
echo Server stopped
pause
