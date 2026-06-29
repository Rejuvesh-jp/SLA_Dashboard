@echo off
:: Check if already running as admin
net session >nul 2>&1
if %errorlevel% neq 0 (
    :: Re-launch this script elevated
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

:: Start PostgreSQL service if not already running
sc query postgresql-x64-18 | findstr /i "RUNNING" >nul 2>&1
if %errorlevel% neq 0 (
    echo Starting PostgreSQL service...
    net start postgresql-x64-18
    timeout /t 3 /nobreak >nul
) else (
    echo PostgreSQL service already running.
)

:: Change to the app directory and start the Python app
cd /d "%~dp0"
echo Starting SLA Dashboard...
python main.py
pause
