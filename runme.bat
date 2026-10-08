@echo off
setlocal enabledelayedexpansion
title Workforce and Turnstile Reports Automation Launcher
cd /d "%~dp0"

echo ===============================================================================
echo        WORKFORCE AND TURNSTILE REPORTS AUTOMATION [POLARS + OPENPYXL]          
echo ===============================================================================
echo.

:: 1. Detect Real Python Executable
set "PYTHON_EXE="

if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
) else if exist "%ProgramFiles%\Python311\python.exe" (
    set "PYTHON_EXE=%ProgramFiles%\Python311\python.exe"
) else if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
) else if exist "%ProgramFiles%\Python312\python.exe" (
    set "PYTHON_EXE=%ProgramFiles%\Python312\python.exe"
)

:: If not found in standard directories, test PATH python (verifying it is not WindowsApps alias)
if not defined PYTHON_EXE (
    python --version >nul 2>&1
    if !ERRORLEVEL! equ 0 (
        set "PYTHON_EXE=python"
    )
)

:: If still not found, test py launcher
if not defined PYTHON_EXE (
    py -3.11 --version >nul 2>&1
    if !ERRORLEVEL! equ 0 (
        set "PYTHON_EXE=py -3.11"
    ) else (
        py --version >nul 2>&1
        if !ERRORLEVEL! equ 0 (
            set "PYTHON_EXE=py"
        )
    )
)

:: If Python is not installed anywhere, install via winget
if not defined PYTHON_EXE (
    echo [!] Python was not detected on your system.
    echo [*] Attempting automated installation of Python 3.11 via winget...
    winget install --id Python.Python.3.11 -e --silent --accept-package-agreements --accept-source-agreements
    if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    ) else (
        echo [ERROR] Python installation failed or requires system restart.
        echo Please install Python 3.11 from https://www.python.org/downloads/
        pause
        exit /b 1
    )
)

echo [*] Using Python: %PYTHON_EXE%

:: 2. Verify / Install Dependencies
"%PYTHON_EXE%" -c "import polars, openpyxl, streamlit" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo.
    echo [*] First-time setup: Installing required dependencies...
    "%PYTHON_EXE%" -m pip install --upgrade pip
    "%PYTHON_EXE%" -m pip install -r requirements.txt
    if %ERRORLEVEL% neq 0 (
        echo [ERROR] Package installation failed. Please check internet connection.
        pause
        exit /b 1
    )
    echo [*] Dependencies successfully verified!
)

:: 3. Launch Streamlit Automation Web UI
echo.
echo [*] Launching Automation Web UI in your default browser...
echo [*] To stop the automation server, close this window or press Ctrl+C
echo.
"%PYTHON_EXE%" -m streamlit run app.py --server.headless=false

pause
