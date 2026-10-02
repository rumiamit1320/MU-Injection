@echo off
setlocal EnableExtensions
cd /d "%~dp0"

title MU Injection Manager - One Click Launcher

echo ============================================================
echo       MU INJECTION MANAGER - ONE CLICK START
echo ============================================================
echo.

REM ------------------------------------------------------------
REM 1. Find Python
REM ------------------------------------------------------------
where py >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON=py"
    goto :python_found
)

where python >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON=python"
    goto :python_found
)

echo [ERROR] Python was not found on this computer.
echo.
echo Install Python 3.11 or newer from:
echo https://www.python.org/downloads/windows/
echo.
echo IMPORTANT: enable "Add Python to PATH" during installation.
pause
exit /b 1

:python_found
echo [OK] Python found.

REM ------------------------------------------------------------
REM 2. Create virtual environment if it does not exist
REM ------------------------------------------------------------
if not exist ".venv\Scripts\python.exe" (
    echo [1/4] Creating Python virtual environment...
    %PYTHON% -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Could not create the virtual environment.
        pause
        exit /b 1
    )
) else (
    echo [1/4] Virtual environment already exists.
)

REM ------------------------------------------------------------
REM 3. Install/update required packages
REM ------------------------------------------------------------
echo [2/4] Installing required packages...
.venv\Scripts\python.exe -m pip install --upgrade pip --disable-pip-version-check
if errorlevel 1 (
    echo [ERROR] pip upgrade failed.
    pause
    exit /b 1
)

.venv\Scripts\python.exe -m pip install -r requirements.txt --disable-pip-version-check
if errorlevel 1 (
    echo [ERROR] Required packages could not be installed.
    pause
    exit /b 1
)

REM ------------------------------------------------------------
REM 4. Start application
REM ------------------------------------------------------------
echo [3/4] Checking application files...
if not exist "app.py" (
    echo [ERROR] app.py was not found.
    pause
    exit /b 1
)
if not exist "MU_Injection_Template.xlsx" (
    echo [ERROR] MU_Injection_Template.xlsx was not found.
    pause
    exit /b 1
)

echo [4/4] Starting MU Injection Manager...
echo.
echo The application will open in your default browser.
echo Keep this window open while using the application.
echo.

.venv\Scripts\python.exe -m streamlit run app.py

if errorlevel 1 (
    echo.
    echo [ERROR] The application stopped unexpectedly.
    pause
)

endlocal
