@echo off
setlocal
cd /d "%~dp0"
title Spectrum Reach - Reporting Ingestion Engine

:: Locate Python 3
set "PY_CMD="
py -3 --version >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD (
    python --version >nul 2>&1 && set "PY_CMD=python"
)
if not defined PY_CMD (
    echo.
    echo   Python 3 was not found on this machine.
    echo   Install Python 3.10 or newer from python.org and select
    echo   "Add python.exe to PATH", or contact IT.
    echo.
    pause
    exit /b 1
)

:: Require Python 3.10+
%PY_CMD% -c "import sys; sys.exit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>&1
if errorlevel 1 (
    echo.
    echo   Python 3.10 or newer is required. Found:
    %PY_CMD% --version
    echo.
    pause
    exit /b 1
)

:: Fully portable folder: app\vendor (populated by developer\tools\
:: make_portable.bat) carries the packages and app\main.py prefers it —
:: skip the install check entirely, nothing needs installing.
if exist "%~dp0app\vendor\" goto :launch

:: Install dependencies only when imports are missing. Preference order:
::   1. Bundled wheelhouse\ folder (ships in the portable release zip) —
::      installs offline, no network needed.
::   2. The internet (pip), as before.
%PY_CMD% -c "import pandas, openpyxl, pptx, tkcalendar, PIL, bs4; import win32com" >nul 2>&1
if errorlevel 1 (
    set "INSTALL_OK="
    if exist "%~dp0wheelhouse\" (
        echo Installing bundled packages ^(offline, first run only^)...
        %PY_CMD% -m pip install --no-index --find-links "%~dp0wheelhouse" -r "%~dp0app\requirements.txt" && set "INSTALL_OK=1"
        if not defined INSTALL_OK echo Bundled packages did not match this Python — trying the internet...
    )
    if not defined INSTALL_OK (
        echo Installing required packages ^(first run only^)...
        %PY_CMD% -m pip install -r "%~dp0app\requirements.txt" && set "INSTALL_OK=1"
    )
    if not defined INSTALL_OK (
        echo.
        echo   Package installation failed. Check the network connection
        echo   or contact IT. The installation error is shown above.
        echo.
        pause
        exit /b 1
    )
    %PY_CMD% -c "import pandas, openpyxl, pptx, tkcalendar, PIL, bs4; import win32com" >nul 2>&1
    if errorlevel 1 (
        echo.
        echo   Packages installed, but one or more still fail to import.
        echo   Send workspace\logs\ingestion_engine.log to support.
        echo.
        pause
        exit /b 1
    )
)

:launch
%PY_CMD% "%~dp0app\main.py"
if errorlevel 1 (
    echo.
    echo   Something went wrong. Review the error above or check:
    echo   %~dp0workspace\logs\ingestion_engine.log
    echo.
    pause
)
endlocal
