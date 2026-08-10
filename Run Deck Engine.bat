@echo off
setlocal
cd /d "%~dp0"
echo Deck Engine is under staged construction.
echo.
echo Stage 0 exposes the headless template-list command only:
echo   py -3 -m app.cli list-templates
echo.
echo The analyst desktop application is scheduled for Stage 5.
pause
exit /b 2
