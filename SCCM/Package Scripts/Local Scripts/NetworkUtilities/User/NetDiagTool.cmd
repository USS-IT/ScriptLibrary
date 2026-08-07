@echo off
SET "_LOGFILE=C:\USS\Logs\User\NetworkUtilities\%~nx0.log"
(
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Show-NetDiagnosticTool.ps1"
) > "%_LOGFILE%" 2>&1