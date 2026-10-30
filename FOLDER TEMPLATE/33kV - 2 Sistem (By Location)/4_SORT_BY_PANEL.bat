@echo off
cd /d "%~dp0"
title BAT 4 - Sort Images by Panel & Equipment
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\sort_by_panel.ps1"
echo.
pause
