@echo off
REM Builds dist\ShiftReportDAR\ShiftReportDAR.exe  (run this ON a Windows PC)
REM Double-click this file, or run it from a Command Prompt in this folder.
setlocal
cd /d "%~dp0"

set "PY=python"
where py >nul 2>nul
if %errorlevel%==0 set "PY=py -3"

%PY% --version >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Install it from https://www.python.org/downloads/
    echo and tick "Add python.exe to PATH" in the installer.
    goto :fail
)

echo [1/4] Creating build environment...
if not exist ".venv-build\Scripts\python.exe" (
    %PY% -m venv .venv-build || goto :fail
)

echo [2/4] Installing openpyxl and PyInstaller...
".venv-build\Scripts\python.exe" -m pip install --upgrade pip >nul
".venv-build\Scripts\python.exe" -m pip install -r requirements.txt pyinstaller || goto :fail

echo [3/4] Building the app...
".venv-build\Scripts\python.exe" -m PyInstaller --noconfirm --clean dar_app.spec || goto :fail

echo [4/4] Adding the editable config and README...
copy /y report_config.json "dist\ShiftReportDAR\" >nul
copy /y README.md "dist\ShiftReportDAR\" >nul

echo.
echo Done!  App: dist\ShiftReportDAR\ShiftReportDAR.exe
echo Copy the whole dist\ShiftReportDAR folder to any Windows PC - no Python needed there.
echo.
pause
exit /b 0

:fail
echo.
echo Build failed - see the messages above.
pause
exit /b 1
