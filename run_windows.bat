@echo off
rem Gamjeong map photo - Windows launcher (double-click). Creates .venv on first run.
chcp 65001 >nul
set PYTHONUTF8=1
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [1/3] Creating virtual environment...
  where py >nul 2>nul
  if %errorlevel%==0 (py -3 -m venv .venv) else (python -m venv .venv)
  if not exist ".venv\Scripts\python.exe" (
    echo Python 3.10+ was not found. Install it from python.org and tick "Add python.exe to PATH".
    pause
    exit /b 1
  )
  echo [2/3] Installing packages...
  ".venv\Scripts\python.exe" -m pip install -q --upgrade pip
  ".venv\Scripts\python.exe" -m pip install -q -r requirements.txt
)

if not exist ".env" copy ".env.example" ".env" >nul

echo [3/3] Starting web app: http://127.0.0.1:8877/
".venv\Scripts\python.exe" -X utf8 webapp_server.py --host 127.0.0.1 --port 8877 --open
pause
