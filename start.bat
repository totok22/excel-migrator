@echo off
REM Windows one-click launcher for Excel Migrator.
setlocal
cd /d "%~dp0"

set VENV_DIR=.venv
set PY=

REM Locate Python 3.10+
for %%C in (py python python3) do (
    %%C -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" 1>/dev/null 2>/dev/null
    if not errorlevel 1 (
        set PY=%%C
        goto :found
    )
)
echo Python 3.10+ not found! Please install: https://www.python.org/downloads/
pause
exit /b 1

:found
if not exist "%VENV_DIR%" (
    echo First run: Creating virtual environment...
    %PY% -m venv %VENV_DIR%
)

call %VENV_DIR%\Scripts\activate.bat

REM Only install dependencies if requirements changed or never installed
set STAMP=%VENV_DIR%\.requirements.stamp
if not exist "%STAMP%" goto :install
fc /b requirements.txt "%STAMP%" >/dev/null 2>/dev/null
if errorlevel 1 goto :install
echo Dependencies ready, skipping installation.
goto :run

:install
echo Installing/updating dependencies...
python -m pip install --upgrade pip >/dev/null
python -m pip install -r requirements.txt >/dev/null
copy /y requirements.txt "%STAMP%" >/dev/null

:run
set PYTHONPATH=%CD%\src;%PYTHONPATH%
python -m excel_migrator.server %*
