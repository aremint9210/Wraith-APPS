@echo off
cd /d "%~dp0"
title BAT 5 - Move Back from Panel Folders
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\move_back_from_panel.ps1"
echo.
pause
