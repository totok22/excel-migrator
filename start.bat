@echo off
REM Windows one-click launcher for Excel Migrator.
setlocal
cd /d "%~dp0"

set "VENV_DIR=.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "PY="

REM Locate Python 3.10+
py -3 -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
if not errorlevel 1 (
    set "PY=py -3"
    goto :found
)

for %%C in (python python3) do (
    %%C -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
    if not errorlevel 1 (
        set "PY=%%C"
        goto :found
    )
)
echo Python 3.10+ not found. Please install it from:
echo https://www.python.org/downloads/
echo.
echo During installation on Windows, enable "Add python.exe to PATH".
pause
exit /b 1

:found
if not exist "%VENV_PY%" (
    echo First run: Creating virtual environment...
    %PY% -m venv "%VENV_DIR%"
    if errorlevel 1 goto :venv_failed
)

if not exist "%VENV_PY%" goto :venv_failed

REM Only install dependencies if requirements changed or never installed
set "STAMP=%VENV_DIR%\.requirements.stamp"
if not exist "%STAMP%" goto :install
fc /b requirements.txt "%STAMP%" >nul 2>nul
if errorlevel 1 goto :install
echo Dependencies ready, skipping installation.
goto :run

:install
echo Installing/updating dependencies...
"%VENV_PY%" -m pip install --upgrade pip >nul
if errorlevel 1 goto :pip_failed
"%VENV_PY%" -m pip install -r requirements.txt >nul
if errorlevel 1 goto :pip_failed
copy /y requirements.txt "%STAMP%" >nul
if errorlevel 1 goto :pip_failed

:run
set "PYTHONPATH=%CD%\src;%PYTHONPATH%"
"%VENV_PY%" -m excel_migrator.server %*
exit /b %ERRORLEVEL%

:venv_failed
echo Failed to create the virtual environment.
echo Please check that Python 3.10+ is installed correctly, then run start.bat again.
pause
exit /b 1

:pip_failed
echo Failed to install dependencies.
echo Please check your network connection, then run start.bat again.
pause
exit /b 1
