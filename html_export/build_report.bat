@echo off
title Amazon Haul EU5 Dashboard Builder

echo.
echo ============================================================
echo   Amazon Haul EU5 Dashboard Builder
echo ============================================================
echo.
echo This will generate a static HTML dashboard from your Excel data.
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8+ and try again.
    pause
    exit /b 1
)

REM Check if input file exists in parent directory
if not exist "..\inputs_forecasting.xlsx" (
    echo ERROR: inputs_forecasting.xlsx not found in parent directory!
    echo Please place your Excel file in the forecasting_dashboard directory.
    pause
    exit /b 1
)

echo Building dashboard...
echo.

python build_dashboard.py --input ..\inputs_forecasting.xlsx --output dashboard_report.html

if errorlevel 1 (
    echo.
    echo Build failed. Check the error messages above.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Dashboard ready!
echo ============================================================
echo.
echo Your dashboard has been saved to: html_export\dashboard_report.html
echo.
echo You can:
echo   - Open it in any browser (double-click the file)
echo   - Share it via email or Slack
echo   - Upload it to Harmony as a static file
echo.
pause
