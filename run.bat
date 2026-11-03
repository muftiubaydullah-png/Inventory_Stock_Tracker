@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (set PY=py) else (set PY=python)
if not exist venv\Scripts\python.exe (
  echo Creating project environment...
  %PY% -m venv venv
  if errorlevel 1 goto failed
)
echo Installing/checking required packages...
venv\Scripts\python.exe -m pip install --upgrade pip
venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Starting Inventory Stock Tracker...
venv\Scripts\python.exe app.py
pause
goto end
:failed
echo Setup failed. Check that Python is installed and internet is available for first-time package installation.
pause
:end
