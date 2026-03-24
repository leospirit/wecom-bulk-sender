@echo off
setlocal
cd /d "%~dp0.."
set "ROOT=%cd%"
set "PORT=18765"
set "PY_CMD=python"
if exist "%ROOT%\.venv\Scripts\python.exe" set "PY_CMD=%ROOT%\.venv\Scripts\python.exe"

:find_free_port
for /f "tokens=1,2,3,4,5" %%a in ('netstat -ano ^| findstr /R /C:":%PORT% .*LISTENING"') do (
  set /a PORT+=1
  goto find_free_port
)

echo [1/3] Checking FastAPI dependencies...
call "%PY_CMD%" -c "import fastapi,uvicorn" >nul 2>nul
if errorlevel 1 (
  echo Installing api requirements...
  call "%PY_CMD%" -m pip install -r api\requirements.txt
  if errorlevel 1 (
    echo Failed to install api requirements.
    pause
    exit /b 1
  )
)

echo [2/3] Starting FastAPI on http://127.0.0.1:%PORT% ...
start "RPA Lite API %PORT%" cmd /k "cd /d %ROOT% && \"%PY_CMD%\" -m uvicorn api.app.main:app --host 127.0.0.1 --port %PORT%"

echo [3/3] Opening browser...
timeout /t 2 /nobreak >nul
start http://127.0.0.1:%PORT%/rpa-lite/

echo Done.
endlocal
