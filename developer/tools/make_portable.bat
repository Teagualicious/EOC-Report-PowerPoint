@echo off
setlocal
cd /d "%~dp0\..\.."
echo Installing dependencies into app\vendor ...
python -m pip install --target app\vendor --upgrade -r app\requirements.txt
if errorlevel 1 (
    echo.
    echo Portable dependency installation failed.
    pause
    exit /b 1
)
echo.
echo Done. app\main.py automatically prefers app\vendor when present.
echo The whole IngestionEngine folder can now be copied to another Windows
echo machine that has Python 3.10+ without reinstalling packages.
pause
endlocal
