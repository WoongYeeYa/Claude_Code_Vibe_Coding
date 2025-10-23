@echo off
echo.
echo ========================================
echo  DOCX to HWP - Build
echo ========================================
echo.

echo Current directory: %CD%
echo.

cd backend-java
if errorlevel 1 (
    echo ERROR: Cannot find backend-java directory!
    echo Make sure you run this from the project root
    pause
    exit /b 1
)

echo Building...
"C:\Program Files\apache-maven-3.9.11\bin\mvn.cmd" clean package -DskipTests

if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR: Build failed!
    echo ========================================
    cd ..
    pause
    exit /b 1
)

cd ..
echo.
echo ========================================
echo  Build Complete!
echo ========================================
echo JAR: backend-java\target\docx-to-hwp-converter-1.0.0.jar
echo.
pause
