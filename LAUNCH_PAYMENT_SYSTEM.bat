@echo off
REM ============================================================
REM Kirembe Secondary School - Payment System Launcher
REM Updated: June 2026 - Includes payroll print fix, Form 3 2026
REM          balance fix, support staff attendance, payroll periods
REM ============================================================

title Kirembe Payment System

echo.
echo ================================================
echo   KIREMBE SECONDARY SCHOOL
echo   Payment Management System
echo ================================================
echo.

REM Change to the folder where this batch file lives
cd /d "%~dp0"

REM Use the exact Python path
set PYTHON_EXE=C:\Users\USER\AppData\Local\Programs\Python\Python315\python.exe

REM Verify Python exists
if not exist "%PYTHON_EXE%" (
    echo ERROR: Python not found at:
    echo %PYTHON_EXE%
    echo.
    pause
    exit /b 1
)

echo Python found: OK
echo.

REM Check if the Excel file exists
if not exist "REMEDIAL_PAYMENT_BALANCES_2026.3_FIXED(1).xlsx" (
    echo ERROR: Excel data file not found.
    echo Looking in: %CD%
    echo.
    pause
    exit /b 1
)

echo Excel file found: OK
echo.

REM Check if server is already running on port 5002
netstat -ano | findstr ":5002" >nul 2>&1
if %errorlevel% equ 0 (
    echo Server already running - opening browser...
    start http://localhost:5002
    timeout /t 3 /nobreak >nul
    exit
)

REM Start the server in a VISIBLE window (not minimized) so errors show
echo Starting payment server...
echo A new window will open - DO NOT CLOSE IT.
echo.
start "Kirembe Payment Server - DO NOT CLOSE" "%PYTHON_EXE%" app_web_new.py

REM Wait for server to be ready
set /a counter=0
echo Waiting for server to load...

:wait_loop
timeout /t 1 /nobreak >nul
netstat -ano | findstr ":5002" >nul 2>&1
if %errorlevel% equ 0 goto server_ready
set /a counter+=1
echo   Loading... %counter% seconds
if %counter% lss 30 goto wait_loop

echo.
echo ERROR: Server did not start in 30 seconds.
echo Check the other window for error messages.
echo.
pause
exit /b 1

:server_ready
timeout /t 2 /nobreak >nul
echo.
echo ================================================
echo   SYSTEM READY! Opening browser...
echo ================================================
echo.
start http://localhost:5002

echo Browser opened to: http://localhost:5002
echo.
echo Keep the "Kirembe Payment Server" window open.
echo Close it when done for the day.
echo.
echo This window closes in 8 seconds...
timeout /t 8 /nobreak >nul
