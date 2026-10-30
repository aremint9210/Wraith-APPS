@echo off
cd /d "%~dp0"
title BAT 5 - Move Images Back from SORT_BY_PANEL
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\move_back_from_panel.ps1"
echo.
pause
