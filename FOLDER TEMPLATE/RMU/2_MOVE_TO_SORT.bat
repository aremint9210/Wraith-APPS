@echo off
cd /d "%~dp0"
title BAT 2 - Move All Images to SORT Folder
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\move_to_sort.ps1"
echo.
pause
