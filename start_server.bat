@echo off
title Process Control Server
cd /d "%~dp0"

REM ================= SETTINGS (edit here) =================
set "GEMINI_MODEL= gemini-2.5-flash-lite"
REM Website access code. Empty = no password.  Example: set "ACCESS_CODE=1234"
set "ACCESS_CODE="
REM PIN before deleting a record. Empty = no PIN.
set "ADMIN_PIN="
REM PIN before generating slides / PIN for editing checklist items
set "SLIDE_PIN=0000"
set "EDIT_PIN=2780"
REM Admin password for the user management page (/users.html)
set "ADMIN_PASS=Admin#12345"
REM Folder for old photos moved by "archive" button. Empty = data\archive_uploads
REM Example (another drive): set "ARCHIVE_DIR=D:\pc-archive"
set "ARCHIVE_DIR="
REM ========================================================

if not exist ".venv\Scripts\python.exe" (
  echo [ERROR] .venv not found. Run first-time setup:
  echo   py -m venv .venv
  echo   .venv\Scripts\python -m pip install -r requirements.txt
  pause
  exit /b 1
)

REM ---- install libraries if missing ----
".venv\Scripts\python.exe" -c "import uvicorn, fastapi, pptx, PIL, multipart" 2>nul
if errorlevel 1 (
  echo Installing libraries - first time only...
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  ".venv\Scripts\python.exe" -c "import uvicorn, fastapi, pptx, PIL, multipart" 2>nul
  if errorlevel 1 (
    echo.
    echo [ERROR] Library install failed - check internet or company proxy, then run again.
    pause
    exit /b 1
  )
)

REM ---- PDF reader for Retort programs (optional: server still runs without it) ----
".venv\Scripts\python.exe" -c "import pdfplumber" 2>nul
if errorlevel 1 (
  echo Installing PDF reader pdfplumber...
  ".venv\Scripts\python.exe" -m pip install pdfplumber
  ".venv\Scripts\python.exe" -c "import pdfplumber" 2>nul
  if errorlevel 1 (
    echo [WARNING] pdfplumber not installed - PDF upload in Retort will not work. Server will still start.
  )
)

REM ---- do not start twice: if port 8000 is already used, stop here ----
netstat -ano | findstr /R /C:":8000 .*LISTENING" >nul
if not errorlevel 1 (
  echo.
  echo  Server is ALREADY RUNNING on port 8000 - in another window.
  echo  No need to start again.
  echo  If the website does not work: close ALL server windows, then run this file again.
  echo.
  pause
  exit /b 0
)

cd app
if "%ACCESS_CODE%"=="" (echo Access code: OFF - no password needed) else (echo Access code: ON)

:loop
netstat -ano | findstr /R /C:":8000 .*LISTENING" >nul
if not errorlevel 1 (
  echo [%date% %time%] Port 8000 is used by another server window - waiting 60 seconds...
  timeout /t 60 >nul
  goto loop
)
echo [%date% %time%] Starting server on port 8000 ...
"..\.venv\Scripts\python.exe" -m uvicorn main:app --host 0.0.0.0 --port 8000
echo [%date% %time%] Server stopped. Restarting in 10 seconds - close this window to stop.
timeout /t 10 >nul
goto loop