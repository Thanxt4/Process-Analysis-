@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title Process Control - Public link

REM ---- team access code (CHANGE THIS) ----
set ACCESS_CODE=1234

REM ---- find Python (python or py) ----
set PY=
if exist ".venv\Scripts\python.exe" set PY=.venv\Scripts\python.exe
if not defined PY (
  where py >nul 2>nul && set PY=py
)
if not defined PY (
  where python >nul 2>nul && set PY=python
)
if not defined PY (
  echo [ERROR] Python not found. Install Python 3.10+ from python.org and tick "Add python.exe to PATH".
  pause & exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo [1/4] Creating .venv ...
  %PY% -m venv .venv || (echo [ERROR] cannot create .venv & pause & exit /b 1)
)
set PY=%~dp0.venv\Scripts\python.exe

echo [2/4] Installing libraries ...
"%PY%" -m pip install -q -r requirements.txt || (echo [ERROR] pip install failed - check internet/proxy & pause & exit /b 1)

echo [3/4] Starting server (new window "ProcessControl-Server") ...
start "ProcessControl-Server" cmd /k "cd /d "%~dp0app" && "%PY%" -m uvicorn main:app --host 0.0.0.0 --port 8000"

if not exist "cloudflared.exe" (
  echo [4/4] Downloading cloudflared.exe ...
  powershell -NoProfile -ExecutionPolicy Bypass -Command "[Net.ServicePointManager]::SecurityProtocol='Tls12'; Invoke-WebRequest -Uri 'https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe' -OutFile 'cloudflared.exe'"
)
if not exist "cloudflared.exe" (
  echo.
  echo [ERROR] Download failed. Download manually in a browser:
  echo   https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe
  echo Rename it to cloudflared.exe and put it in this folder: %~dp0
  echo The server is still running - open http://localhost:8000 on this PC.
  pause & exit /b 1
)

timeout /t 5 >nul
echo.
echo ============================================================
echo  Public link = the line  https://....trycloudflare.com  below
echo  Access code: %ACCESS_CODE%      (keep this window open)
echo ============================================================
echo.
cloudflared.exe tunnel --no-autoupdate --url http://localhost:8000
echo.
echo cloudflared stopped.
pause
