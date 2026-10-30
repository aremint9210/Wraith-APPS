@echo off
cd /d "%~dp0"
title BAT 1 - Rename Images (Master List)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\rename_images.ps1"
echo.
pause
