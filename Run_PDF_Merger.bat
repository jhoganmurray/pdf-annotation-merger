@echo off
title PDF Annotation Merger

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
    echo.
    pip install pymupdf
    if errorlevel 1 (
        echo.
        echo Failed to install PyMuPDF. Please run this command manually:
        echo pip install pymupdf
        echo.
        pause
        exit /b 1
    )
    echo.
    echo Installation complete!
    echo.
)

REM Run the application
echo Starting PDF Annotation Merger...
pythonw "%~dp0PDF_Annotation_Merger.pyw"
