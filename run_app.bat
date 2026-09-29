@echo off
title OmniStudy Academic Workspace
echo ========================================================
echo        OmniStudy - Starting on New Device
echo ========================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH!
    echo Please install Python 3.10+ from https://python.org and check "Add Python to PATH".
    pause
    exit /b 1
)

:: Set up virtual environment if not present
if not exist ".venv\Scripts\activate.bat" (
    echo [*] Creating virtual environment (.venv)...
    python -m venv .venv
    echo [*] Activating virtual environment...
    call .venv\Scripts\activate.bat
    echo [*] Installing dependencies from requirements.txt...
    python -m pip install --upgrade pip
    pip install -r requirements.txt
) else (
    call .venv\Scripts\activate.bat
)

echo.
echo [*] Launching OmniStudy Server...
echo [*] Open your browser at: http://127.0.0.1:5000
echo.
python app.py
pause
