@echo off
title ECU-TEST PRJ Matrix Validator
cd /d "%~dp0"
python app.py
if errorlevel 1 (
    echo.
    echo An error occurred while launching the validator.
    pause
)
