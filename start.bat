@echo off
REM -- Rebate Form Generator - local web UI launcher -----------------------
REM Starts the local server (127.0.0.1) and opens it in Chrome app mode.
REM Closing the app window stops the server automatically.

title Rebate Form Generator
cd /d "%~dp0"

set "PYTHONPATH=%~dp0src"
set "PY=%~dp0.venv\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

"%PY%" -m rebate_form_generator.main %*

if errorlevel 1 (
    echo.
    echo [ERROR] The server exited unexpectedly.
    echo         Check that Python is installed and that .venv exists.
    echo.
    pause
)
