@Echo off
setlocal EnableDelayedExpansion
REM General Ivanti / Pulse Secure VPN Config
REM aka Default
SET "_LOGFILE=C:\USS\Logs\Packages\SwitchVPN\general.cmd.log"
(
	SET "_PULSECONFIG=General.pulsepreconfig"
	If Not Exist C:\temp MD C:\temp
	ECHO Copying [%~dp0!_PULSECONFIG!] to C:\temp
	Copy "%~dp0%_PULSECONFIG%" c:\temp /y
	CD "%ProgramFiles(x86)%\Common Files\Pulse Secure\JamUI"
	echo Importing [c:\temp\!_PULSECONFIG!]
	jamCommand.exe /importfile "c:\temp\!_PULSECONFIG!"
) > "%_LOGFILE%" 2>&1
