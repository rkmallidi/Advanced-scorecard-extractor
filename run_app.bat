@echo off
setlocal

cd /d "%~dp0"

echo ========================================
echo JEE Advanced Scorecard Extractor
echo ========================================
echo.

set "PYTHON_CMD="

where python >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=py"
    )
)

if "%PYTHON_CMD%"=="" (
    echo Python was not found.
    echo Please install Python 3.10 or newer and try again.
    pause
    exit /b 1
)

echo Using Python:
%PYTHON_CMD% --version
echo.

echo Upgrading pip...
%PYTHON_CMD% -m pip install --upgrade pip
if errorlevel 1 (
    echo.
    echo Failed to upgrade pip.
    pause
    exit /b 1
)

echo.
echo Installing/verifying required Python packages...
%PYTHON_CMD% -m pip install pandas openpyxl beautifulsoup4 playwright
if errorlevel 1 (
    echo.
    echo Failed to install required packages.
    pause
    exit /b 1
)

echo.
echo Installing/verifying Playwright Chromium browser...
%PYTHON_CMD% -m playwright install chromium
if errorlevel 1 (
    echo.
    echo Failed to install Playwright Chromium.
    pause
    exit /b 1
)

echo.
echo Starting application...
%PYTHON_CMD% src\jee_response_extractor_ui.py

echo.
echo Application closed.
pause
