@echo off
REM Word Document Checker - Startup Script
REM ========================================

echo.
echo Starting Word Document Checker...
echo.

REM Add src to PYTHONPATH
set PYTHONPATH=%~dp0src
cd /d "%~dp0src"

REM Check Python installation
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found!
    echo.
    echo Please install Python 3.8 or higher from:
    echo https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

REM Check dependencies
echo Checking dependencies...
python -c "import PyQt5" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing dependencies...
    echo.
    cd /d "%~dp0"
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [ERROR] Failed to install dependencies!
        echo.
        echo Try running manually:
        echo   python -m pip install PyQt5 python-docx openpyxl
        echo.
        pause
        exit /b 1
    )
    cd /d "%~dp0src"
    echo.
    echo [OK] Dependencies installed.
    echo.
)

REM Launch application
echo Launching application...
echo.
python main.py

REM Keep window open if there's an error
if errorlevel 1 (
    echo.
    echo [ERROR] Application failed to start!
    echo.
    pause
)
