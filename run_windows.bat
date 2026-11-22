@echo off
REM Runs the app from source on Windows (no packaging). First run sets things up.
setlocal
cd /d "%~dp0"

set "PY=python"
where py >nul 2>nul
if %errorlevel%==0 set "PY=py -3"

if not exist ".venv\Scripts\pythonw.exe" (
    echo First run: setting up...
    %PY% -m venv .venv || goto :fail
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt || goto :fail
)
start "" ".venv\Scripts\pythonw.exe" dar_app.py %*
exit /b 0

:fail
echo.
echo Setup failed. Is Python installed (python.org, with "Add python.exe to PATH" ticked)?
pause
exit /b 1
