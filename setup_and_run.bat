@echo off
setlocal EnableDelayedExpansion

REM ============================================================
REM  Amazon Haul EU5 Forecasting Dashboard - One-Click Setup
REM ============================================================
REM  This script will:
REM  1. Check if Python is installed
REM  2. Create a virtual environment (if not exists)
REM  3. Install required packages
REM  4. Launch the dashboard
REM ============================================================

title Amazon Haul EU5 Forecasting Dashboard Setup

echo.
echo ============================================================
echo   Amazon Haul EU5 Forecasting Dashboard - Setup
echo ============================================================
echo.

REM Check if Python is installed
echo [1/4] Checking for Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Python is not installed or not in PATH!
    echo.
    echo Please install Python 3.8 or later from:
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANT: During installation, check the box
    echo   "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo    Found Python %PYTHON_VERSION%
echo.

REM Check if virtual environment exists
echo [2/4] Setting up virtual environment...
if not exist "venv" (
    echo    Creating new virtual environment...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo ERROR: Failed to create virtual environment!
        pause
        exit /b 1
    )
    echo    Virtual environment created.
) else (
    echo    Virtual environment already exists.
)
echo.

REM Activate virtual environment
call venv\Scripts\activate.bat
if %errorlevel% neq 0 (
    echo ERROR: Failed to activate virtual environment!
    pause
    exit /b 1
)

REM Install/update requirements
echo [3/4] Installing required packages...
echo    This may take a few minutes on first run...
pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo.
    echo WARNING: Some packages may have failed to install.
    echo    Trying again with verbose output...
    pip install -r requirements.txt
)
echo    Packages installed successfully.
echo.

REM Launch the dashboard
echo [4/4] Launching dashboard...
echo.
echo ============================================================
echo   Dashboard starting at: http://localhost:5000
echo ============================================================
echo.
echo   Opening browser in 3 seconds...
echo   (Press Ctrl+C to stop the server)
echo.

REM Open browser after a delay
start /b cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:5000"

REM Start the Flask app
python app.py

REM Keep window open if server stops
pause
