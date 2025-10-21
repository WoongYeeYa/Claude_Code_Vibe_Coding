@echo off
echo ==========================================
echo DOCX to HWP Converter - Run Server
echo ==========================================
echo.

REM Set Java and Maven paths
set "JAVA_HOME=C:\Program Files\Java\jdk-17"
set "MAVEN_HOME=C:\Program Files\apache-maven-3.9.11"
set "PATH=%JAVA_HOME%\bin;%MAVEN_HOME%\bin;%PATH%"

echo Starting Spring Boot application...
echo Server will be available at: http://localhost:8000
echo.

call mvn spring-boot:run

pause
