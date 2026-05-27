@echo off
REM Windows one-click launcher for Excel Migrator.
setlocal
cd /d "%~dp0"

set VENV_DIR=.venv
set PY=

REM Locate Python 3.10+
for %%C in (py python python3) do (
    %%C -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" 1>nul 2>nul
    if not errorlevel 1 (
        set PY=%%C
        goto :found
    )
)
echo 未找到 Python 3.10+，请先安装：https://www.python.org/downloads/
pause
exit /b 1

:found
if not exist "%VENV_DIR%" (
    echo 首次运行：建立虚拟环境...
    %PY% -m venv %VENV_DIR%
)

call %VENV_DIR%\Scripts\activate.bat

REM Only install dependencies if requirements changed or never installed
set STAMP=%VENV_DIR%\.requirements.stamp
if not exist "%STAMP%" goto :install
fc /b requirements.txt "%STAMP%" >nul 2>nul
if errorlevel 1 goto :install
echo 依赖已就绪，跳过安装。
goto :run

:install
echo 安装/更新依赖...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt >nul
copy /y requirements.txt "%STAMP%" >nul

:run
set PYTHONPATH=%CD%\src;%PYTHONPATH%
python -m excel_migrator.server %*
