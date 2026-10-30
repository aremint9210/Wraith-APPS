@echo off
cd /d "%~dp0"
title BAT 3 - Restore Images from SORT
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\move_back.ps1"
echo.
pause
