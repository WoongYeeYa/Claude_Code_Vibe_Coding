@echo off
echo ==========================================
echo DOCX to HWP Converter - Build Script
echo ==========================================
echo.

REM Set Java and Maven paths
set "JAVA_HOME=C:\Program Files\Java\jdk-17"
set "MAVEN_HOME=C:\Program Files\apache-maven-3.9.11"
set "PATH=%JAVA_HOME%\bin;%MAVEN_HOME%\bin;%PATH%"

echo Building project with Maven...
call mvn clean package -DskipTests

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ==========================================
    echo Build completed successfully!
    echo JAR file location: target\docx-to-hwp-converter-1.0.0.jar
    echo ==========================================
) else (
    echo.
    echo ==========================================
    echo Build failed! Please check the error messages above.
    echo ==========================================
)

pause
