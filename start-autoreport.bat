@echo off
setlocal

set "ROOT=%~dp0"
set "BACKEND_DIR=%ROOT%backend"
set "PYTHON=%BACKEND_DIR%\.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
  echo [ERROR] Python virtual environment not found.
  echo Expected: %PYTHON%
  pause
  exit /b 1
)

start "AutoReport Backend" cmd /k ""cd /d "%BACKEND_DIR%" && "%PYTHON%" -m uvicorn main:app --host 0.0.0.0 --port 8000""
timeout /t 3 /nobreak >nul
start "AutoReport Tunnel" cmd /k ""cd /d "%ROOT%" && npx -y localtunnel --port 8000 --subdomain autoreport-backend""

echo.
echo AutoReport backend and tunnel are starting.
echo Backend: http://127.0.0.1:8000
echo Tunnel:  https://autoreport-backend.loca.lt
echo.
echo Keep both windows open while the site is in use.
pause