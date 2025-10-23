@echo off
echo.
echo ========================================
echo  DOCX to HWP - Install
echo ========================================
echo.

echo Current directory: %CD%
echo.

echo [1/2] Building hwplib...
cd hwplib-main
if errorlevel 1 (
    echo ERROR: Cannot find hwplib-main directory!
    pause
    exit /b 1
)

"C:\Program Files\apache-maven-3.9.11\bin\mvn.cmd" clean install -DskipTests
if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR: hwplib build failed!
    echo ========================================
    cd ..
    pause
    exit /b 1
)
echo SUCCESS: hwplib built
echo.

cd ..
echo [2/2] Building backend...
cd backend-java
if errorlevel 1 (
    echo ERROR: Cannot find backend-java directory!
    pause
    exit /b 1
)

"C:\Program Files\apache-maven-3.9.11\bin\mvn.cmd" clean package -DskipTests
if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR: backend build failed!
    echo ========================================
    cd ..
    pause
    exit /b 1
)
echo SUCCESS: backend built
echo.

cd ..
echo.
echo ========================================
echo  Installation Complete!
echo ========================================
echo.
echo Next: Run start_server.bat
echo.
pause
