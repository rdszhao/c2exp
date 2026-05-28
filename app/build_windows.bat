@echo off
REM Build script for Concept2 Exporter Windows executable
REM Run this on a Windows machine with Python 3.10+ installed

echo ============================================
echo Building Concept2 Logbook Exporter
echo ============================================

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.10+ from https://python.org
    pause
    exit /b 1
)

REM Create virtual environment if it doesn't exist
if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Install dependencies
echo Installing dependencies...
pip install -r requirements.txt

REM Build executable
echo Building executable...
pyinstaller concept2_export.spec --clean

echo.
echo ============================================
echo Build complete!
echo ============================================
echo.
echo The executable is at: dist\Concept2Exporter.exe
echo.
echo You can distribute this single .exe file to users.
echo No installation required - just run the exe!
echo.

pause
