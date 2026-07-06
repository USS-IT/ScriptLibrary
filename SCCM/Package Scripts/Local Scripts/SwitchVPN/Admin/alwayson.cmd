@Echo off
setlocal EnableDelayedExpansion
REM Always On Ivanti / Pulse Secure VPN Config
REM aka Premium
SET "_LOGFILE=C:\USS\Logs\Packages\SwitchVPN\alwayson.cmd.log"
(
	SET "_PULSECONFIG=AlwaysOn.pulsepreconfig"
	If Not Exist C:\temp MD C:\temp
	ECHO Copying [%~dp0!_PULSECONFIG!] to C:\temp
	Copy "%~dp0%_PULSECONFIG%" c:\temp /y
	CD "%ProgramFiles(x86)%\Common Files\Pulse Secure\JamUI"
	echo Importing [c:\temp\!_PULSECONFIG!]
	jamCommand.exe /importfile "c:\temp\!_PULSECONFIG!"
) > "%_LOGFILE%" 2>&1