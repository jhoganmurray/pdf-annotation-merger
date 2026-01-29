@echo off
title PDF Comment Collector

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Python is not installed or not in PATH.
    echo.
    echo Please install Python from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

REM Check if PyMuPDF is installed, install if not
python -c "import fitz" >nul 2>&1
if errorlevel 1 (
    echo Installing required dependency (PyMuPDF)...
    pip install pymupdf
    if errorlevel 1 (
        echo Failed to install. Please run: pip install pymupdf
        pause
        exit /b 1
    )
)

REM Check if pikepdf is installed, install if not (for direct merge feature)
python -c "import pikepdf" >nul 2>&1
if errorlevel 1 (
    echo Installing optional dependency (pikepdf) for direct PDF merge...
    pip install pikepdf
    if errorlevel 1 (
        echo Note: pikepdf not installed. Merge feature will be disabled.
        echo You can still use XFDF export.
    )
)

REM Run the application
pythonw "%~dp0PDF_Comment_Collector.pyw"
