@echo off
setlocal
cd /d "%~dp0\.."
py -3 -m app.mapper
if errorlevel 1 pause
