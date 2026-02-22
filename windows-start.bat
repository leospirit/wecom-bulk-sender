@echo off
setlocal

where docker >nul 2>&1
if errorlevel 1 (
  echo Docker Desktop not found. Please install Docker Desktop and enable WSL2.
  pause
  exit /b 1
)

docker info >nul 2>&1
if errorlevel 1 (
  echo Docker is not running. Please start Docker Desktop and wait until it says Running.
  pause
  exit /b 1
)

echo Starting containers...
cd /d %~dp0

docker compose up --build -d
if errorlevel 1 (
  echo Failed to start containers.
  pause
  exit /b 1
)

echo Opening UI...
start http://localhost:5173

endlocal
