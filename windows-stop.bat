@echo off
setlocal

where docker >nul 2>&1
if errorlevel 1 (
  echo Docker Desktop not found.
  pause
  exit /b 1
)

docker info >nul 2>&1
if errorlevel 1 (
  echo Docker is not running.
  pause
  exit /b 1
)

echo Stopping containers...
cd /d %~dp0

docker compose down
if errorlevel 1 (
  echo Failed to stop containers.
  pause
  exit /b 1
)

echo Done.

endlocal
