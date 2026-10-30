@echo off
cd /d "%~dp0"
title BAT 1 - Rename Images (VCB By Location)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0_engine\rename_images.ps1"
echo.
pause
