@echo off
pushd "%~dp0"
powershell.exe -NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File "Invoke-DCUUpdateCheck.ps1"
popd
exit /b
