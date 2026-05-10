@echo off
echo ========================================
echo   poinmeet - Installation
echo ========================================
echo.

cd /d "%~dp0backend"

echo [1/3] Creating virtual environment...
if exist venv (
    echo  [*] Removing old venv...
    rmdir /s /q venv
)
python -m venv venv
if errorlevel 1 (
    echo [ERROR] Python not found. Install Python 3.10+ from https://python.org
    pause
    exit /b
)

echo [2/3] Activating virtual environment...
call venv\Scripts\activate

echo [3/3] Installing dependencies...
pip install -r requirements.txt

echo.
echo ========================================
echo   Installation complete!
echo   Launching app...
echo ========================================
echo.
cd /d "%~dp0"
call start.bat
