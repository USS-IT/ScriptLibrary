@echo off
pushd "%~dp0"
powershell.exe -NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File "Invoke-DCUUpdateCheck.ps1" -UpdateType "driver,application,others"
popd
exit /b
