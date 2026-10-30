@echo off
cd /d "%~dp0"
title BAT 3 - Move Images from SORT Back to Folders
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\move_back.ps1"
echo.
pause
