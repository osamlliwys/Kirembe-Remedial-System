@echo off
echo ================================================
echo Kirembe Secondary School
echo Payment Management System - Web App
echo ================================================
echo.
echo Starting web server...
echo.
echo Once started, open your browser and go to:
echo   http://localhost:5002
echo.
echo Press Ctrl+C to stop the server
echo.
echo ================================================
echo.

py app_web_new.py

if errorlevel 1 (
    echo.
    echo ================================================
    echo ERROR: Failed to start web server
    echo ================================================
    echo.
    echo Possible solutions:
    echo 1. Install Python from https://python.org
    echo 2. Install dependencies: pip install -r requirements.txt
    echo 3. Check that the Excel file exists
    echo 4. Make sure port 5002 is not in use
    echo.
    pause
)
