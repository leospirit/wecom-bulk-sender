@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo [1/5] Detecting Python...
set "PY_CMD="
where python >nul 2>nul && set "PY_CMD=python"
if "!PY_CMD!"=="" (
  where py >nul 2>nul && set "PY_CMD=py -3"
)
if "!PY_CMD!"=="" (
  echo Python 3 not found. Please install Python 3.11+ first.
  pause
  exit /b 1
)

set "VENV_DIR=.venv_rpa"
if not exist "%VENV_DIR%\Scripts\python.exe" (
  echo [2/5] Creating virtual environment...
  call !PY_CMD! -m venv "%VENV_DIR%"
)

set "RUN_CMD=!PY_CMD!"
if exist "%VENV_DIR%\Scripts\python.exe" (
  set "RUN_CMD="%VENV_DIR%\Scripts\python.exe""
  echo [3/5] Using virtual environment Python.
) else (
  echo [3/5] Virtual environment unavailable. Using system Python.
)

echo [4/5] Installing dependencies...
call !RUN_CMD! -m ensurepip --upgrade >nul 2>nul
call !RUN_CMD! -m pip --version >nul 2>nul
if errorlevel 1 (
  if not "!RUN_CMD!"=="!PY_CMD!" (
    echo venv pip unavailable. Falling back to system Python pip.
    set "RUN_CMD=!PY_CMD!"
    call !RUN_CMD! -m pip --version >nul 2>nul
  )
)
if errorlevel 1 (
  echo pip is unavailable. Please install Python with pip enabled.
  pause
  exit /b 1
)

call !RUN_CMD! -m pip install --upgrade pip >nul
call !RUN_CMD! -m pip install -r tools\rpa-requirements.txt
if errorlevel 1 (
  echo Dependency installation failed.
  pause
  exit /b 1
)

echo [5/5] Launching RPA UI...
call !RUN_CMD! tools\wecom_rpa_gui.py

endlocal
