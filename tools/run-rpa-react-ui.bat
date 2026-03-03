@echo off
setlocal

set ROOT=%~dp0..
cd /d "%ROOT%"

echo [1/3] Starting FastAPI on http://localhost:8000 ...
start "RPA API 8000" cmd /k "cd /d %ROOT% && python -m uvicorn api.app.main:app --host 0.0.0.0 --port 8000"

echo [2/3] Starting React UI on http://localhost:5173 ...
start "RPA Web 5173" cmd /k "cd /d %ROOT%\web && set VITE_API_BASE=http://localhost:8000 && npm run dev -- --host 0.0.0.0 --port 5173"

echo [3/3] Opening browser...
timeout /t 2 /nobreak >nul
start http://localhost:5173

echo Done.
