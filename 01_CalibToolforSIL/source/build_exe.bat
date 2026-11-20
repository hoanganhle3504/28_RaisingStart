@echo off
title Build SiL_CalibTool.exe
cd /d "%~dp0"
echo Rebuilding standalone executable from source...
pyinstaller --noconfirm --onefile --windowed --name SiL_CalibTool main.py
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build successful! Moving executable to root directory...
    move /y dist\SiL_CalibTool.exe ..\SiL_CalibTool.exe
    rd /s /q build dist __pycache__ 2>nul
    del /f /q SiL_CalibTool.spec 2>nul
    echo Done! Output: C:\Workspace\25_CalibToolforSIL\SiL_CalibTool.exe
) else (
    echo.
    echo Build failed with error %ERRORLEVEL%.
)
pause
