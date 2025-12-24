@echo off
title Amazon Haul EU5 Forecasting Dashboard
color 0A

echo.
echo ============================================================
echo    Amazon Haul EU5 Forecasting Dashboard
echo ============================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please install Python 3.8+ and try again.
    pause
    exit /b 1
)

:: Navigate to script directory
cd /d "%~dp0"

:: Check if virtual environment exists
if not exist "venv" (
    echo [INFO] Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
)

:: Activate virtual environment
echo [INFO] Activating virtual environment...
call venv\Scripts\activate.bat

:: Install/update dependencies
echo [INFO] Installing dependencies...
pip install -r requirements.txt --quiet

:: Check if installation was successful
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    echo Please check requirements.txt and try again.
    pause
    exit /b 1
)

echo.
echo [SUCCESS] Dependencies installed successfully!
echo.
echo ============================================================
echo    Starting Dashboard Server...
echo ============================================================
echo.
echo    Dashboard URL: http://localhost:5000
echo    Press Ctrl+C to stop the server
echo.
echo ============================================================
echo.

:: Wait a moment then open browser
start "" cmd /c "timeout /t 2 /nobreak >nul && start http://localhost:5000"

:: Run the Flask application
python app.py

:: Deactivate virtual environment on exit
call venv\Scripts\deactivate.bat

pause
