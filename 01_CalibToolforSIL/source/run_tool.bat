@echo off
title SiL Software Calibration Tool
cd /d "%~dp0"
python main.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Application terminated with error %ERRORLEVEL%.
    pause
)
