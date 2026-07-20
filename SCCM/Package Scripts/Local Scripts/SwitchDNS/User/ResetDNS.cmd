@echo off
SET "_LOGFILE=C:\USS\Logs\User\SwitchDNS\%~nx0.log"
(
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Install-Tasks-User.ps1"
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -Command "Write-EventLog -LogName 'USS-EventLog' -Source 'SwitchDNS-Reset' -EventID 1000 -EntryType Information -Message 'Trigger Switch-DNS -Mode Reset'"
) > "%_LOGFILE%" 2>&1