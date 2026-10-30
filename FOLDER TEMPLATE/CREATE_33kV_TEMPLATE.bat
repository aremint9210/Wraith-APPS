@echo off
cd /d "%~dp0"
title 33kV Substation Folder Template Generator
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0create_33kv_template.ps1"
echo.
pause
