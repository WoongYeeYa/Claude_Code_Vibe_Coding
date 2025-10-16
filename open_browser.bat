@echo off
chcp 65001 >nul
echo Opening Word to HWP Converter in browser...
timeout /t 1 /nobreak >nul
start http://localhost:8000
