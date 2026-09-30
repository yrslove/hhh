@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Doctor Eleks Automation

echo ==============================================
echo        Doctor Eleks Automation
echo ==============================================
echo.
echo The application will open in your browser.
echo Keep this window open while it is running.
echo.

where py >nul 2>&1
if not errorlevel 1 (
  set "PY_CMD=py -3"
) else (
  where python >nul 2>&1
  if errorlevel 1 (
    echo ERROR: Python 3.11 or newer was not found.
    echo Install Python and run this file again.
    goto :error
  )
  set "PY_CMD=python"
)

%PY_CMD% -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>&1
if errorlevel 1 (
  echo ERROR: Python 3.11 or newer is required.
  %PY_CMD% --version
  goto :error
)

if not exist ".venv\Scripts\python.exe" goto :create_venv
".venv\Scripts\python.exe" --version >nul 2>&1
if not errorlevel 1 (
  echo [1/3] Local Python environment found.
  goto :check_dependencies
)

echo [1/3] Repairing incomplete local Python environment...
rmdir /s /q ".venv" >nul 2>&1
if exist ".venv\Scripts\python.exe" (
  echo ERROR: The local Python environment is in use and cannot be repaired.
  echo Close other application windows and run this file again.
  goto :error
)

:create_venv
echo [1/3] Creating local Python environment...
%PY_CMD% -m venv .venv
if errorlevel 1 goto :error

:check_dependencies

echo [2/3] Checking dependencies...
".venv\Scripts\python.exe" -c "import fastapi,uvicorn,openpyxl,multipart,playwright,httpx" >nul 2>&1
if errorlevel 1 (
  echo       Installing dependencies...
  ".venv\Scripts\python.exe" -m pip install --disable-pip-version-check -q -r requirements.txt
  if errorlevel 1 goto :error
)

echo [3/3] Starting application...
".venv\Scripts\python.exe" scripts\launcher.py --reload
if errorlevel 1 goto :error
exit /b 0

:error
echo.
echo The application could not be started. See the message above.
pause
exit /b 1
