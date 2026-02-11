@echo off
pushd "%~dp0"
powershell.exe -NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File "UninstallScriptPS.ps1" -SoftwareName "ActivID ActivClient"
popd
exit /b