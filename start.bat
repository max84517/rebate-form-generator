@echo off
REM -- Rebate Form Generator - local web UI launcher -----------------------
REM First run: creates .venv and installs the two dependencies.
REM After that: starts the local server and opens it in Chrome app mode.
REM Closing the app window stops the server.

title Rebate Form Generator
cd /d "%~dp0"

set "VENV=%~dp0.venv"
set "PY=%VENV%\Scripts\python.exe"

if not exist "%PY%" (
    echo First run - setting up the Python environment.
    echo This takes a minute and only happens once.
    echo.

    where py >nul 2>nul
    if errorlevel 1 (
        set "BOOTSTRAP=python"
    ) else (
        set "BOOTSTRAP=py -3"
    )

    %BOOTSTRAP% -m venv "%VENV%"
    if not exist "%PY%" (
        echo.
        echo [ERROR] Could not create the virtual environment.
        echo         Install Python 3.10 or newer from python.org or the
        echo         Microsoft Store, then run this file again.
        echo.
        pause
        exit /b 1
    )

    "%PY%" -m pip install --upgrade pip --quiet
    "%PY%" -m pip install --quiet "openpyxl>=3.1" "python-docx>=1.1"
    if errorlevel 1 (
        echo.
        echo [ERROR] Could not install the dependencies.
        echo         Check your network or proxy settings and try again.
        echo.
        pause
        exit /b 1
    )
    echo Setup complete.
    echo.
)

set "PYTHONPATH=%~dp0src"
"%PY%" -m rebate_form_generator.main %*

if errorlevel 1 (
    echo.
    echo [ERROR] The server exited unexpectedly.
    echo.
    pause
)
