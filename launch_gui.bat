@echo off
setlocal
cd /d "%~dp0"

title Indeed Proto QA Reviewer - Launcher

echo ============================================================
echo   Indeed Proto QA Reviewer - Launch Script
echo ============================================================
echo.

:: ── Python check ─────────────────────────────────────────────────────────────
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo         Please install Python 3.10+ and re-run this launcher.
    echo         Download: https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [INFO] Python found:
python --version
echo.

:: ── pip upgrade ───────────────────────────────────────────────────────────────
echo [INFO] Upgrading pip...
python -m pip install --upgrade pip --quiet
if %errorlevel% neq 0 (
    echo [WARN] Could not upgrade pip - continuing anyway...
    echo.
)

:: ── requirements ─────────────────────────────────────────────────────────────
if not exist requirements.txt (
    echo [ERROR] requirements.txt not found in:
    echo         %~dp0
    echo         Please ensure requirements.txt is in the same folder as launch.bat
    echo.
    pause
    exit /b 1
)

echo [INFO] Installing / verifying requirements...
python -m pip install -r requirements.txt --quiet
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to install one or more requirements.
    echo         Try running this script as Administrator, or check your
    echo         internet connection and requirements.txt contents.
    echo.
    pause
    exit /b 1
)

echo [INFO] All requirements satisfied.
echo.

:: ── launch app ───────────────────────────────────────────────────────────────
if not exist app.py (
    echo [ERROR] app.py not found in:
    echo         %~dp0
    echo         Please ensure app.py is in the same folder as launch.bat
    echo.
    pause
    exit /b 1
)

echo [INFO] Launching Indeed Proto QA Reviewer...
echo ============================================================
echo.

python app.py

:: ── exit handling ─────────────────────────────────────────────────────────────
if %errorlevel% neq 0 (
    echo.
    echo ============================================================
    echo [ERROR] Application exited with an error (code: %errorlevel%)
    echo         Check the output above for details.
    echo ============================================================
    echo.
    pause
    exit /b 1
)

echo.
echo [INFO] Application closed normally.
endlocal