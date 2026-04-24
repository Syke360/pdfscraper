@echo off
pushd "%~dp0"
echo Checking script location...
if not exist "Scripts\PDFPicker2.ps1" (
    echo ERROR: Scripts\PDFPicker2.ps1 not found!
    pause
    exit
)
echo Launching PowerShell...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "Scripts\PDFPicker2.ps1"
if %errorlevel% neq 0 pause
popd