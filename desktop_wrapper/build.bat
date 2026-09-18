@echo off
setlocal

cd /d "%~dp0"
set "VENV_DIR=%~dp0.venv"
set "PYTHON=%VENV_DIR%\Scripts\python.exe"

if not exist "%PYTHON%" (
    echo Creating the desktop wrapper virtual environment...
    py -3 -m venv "%VENV_DIR%"
    if errorlevel 1 goto :error
)

echo Installing build dependencies...
"%PYTHON%" -m pip install --upgrade pip
if errorlevel 1 goto :error
"%PYTHON%" -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo Removing old build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "EDP Analytics.spec" del /q "EDP Analytics.spec"

echo Building EDP Analytics.exe...
"%PYTHON%" -m PyInstaller --noconfirm --clean --onefile --windowed --name "EDP Analytics" desktop_wrapper.py
if errorlevel 1 goto :error

echo.
echo Build complete: "%~dp0dist\EDP Analytics.exe"
exit /b 0

:error
echo.
echo Build failed with exit code %errorlevel%.
exit /b %errorlevel%