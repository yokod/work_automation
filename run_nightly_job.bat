@echo off
chcp 65001 >nul
cd /d "d:\yoel\projects\auto"
for /f "usebackq tokens=*" %%i in (`powershell -NoProfile -Command "(Get-Date).ToString('yyyy-MM-dd')"`) do set LOG_DATE=%%i
set LOG_FILE=logs\nightly_%LOG_DATE%.log
set PYTHONIOENCODING=utf-8
echo [%DATE% %TIME%] Starting Nightly Job... >> "%LOG_FILE%" 2>&1
python run_nightly_job.py >> "%LOG_FILE%" 2>&1
echo [%DATE% %TIME%] Job Complete. >> "%LOG_FILE%" 2>&1
