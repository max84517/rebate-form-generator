@echo off
REM -- Rebate Form Generator - local web UI launcher -----------------------
REM First run: creates .venv and installs the two dependencies.
REM After that: starts the local server and opens it in Chrome app mode.
REM Closing the app window stops the server.
REM
REM Note: no "set" inside a parenthesised block. Batch expands %VAR% when it
REM parses the whole block, so a value assigned inside it reads back empty.

title Rebate Form Generator
cd /d "%~dp0"

set "VENV=%~dp0.venv"
set "PY=%VENV%\Scripts\python.exe"

if exist "%PY%" goto :run

echo.
echo First run - setting up the Python environment.
echo This takes a minute and only happens once.
echo.

py -3 -m venv "%VENV%" >nul 2>nul
if exist "%PY%" goto :deps
python -m venv "%VENV%" >nul 2>nul
if exist "%PY%" goto :deps

echo [ERROR] Could not create the virtual environment.
echo         Install Python 3.10 or newer from python.org or the
echo         Microsoft Store, then run this file again.
echo.
pause
exit /b 1

:deps
echo Installing openpyxl and python-docx...
"%PY%" -m pip install --upgrade pip --quiet
"%PY%" -m pip install --quiet "openpyxl>=3.1" "python-docx>=1.1"
if errorlevel 1 goto :depsfailed
echo Setup complete.
echo.
goto :run

:depsfailed
echo.
echo [ERROR] Could not install the dependencies.
echo         Check your network or proxy settings and try again.
echo.
pause
exit /b 1

:run
set "PYTHONPATH=%~dp0src"
"%PY%" -m rebate_form_generator.main %*

if errorlevel 1 (
    echo.
    echo [ERROR] The server exited unexpectedly.
    echo.
    pause
)
