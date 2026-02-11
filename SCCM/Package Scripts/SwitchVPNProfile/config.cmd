@Echo off
:: Switch between General and Always ON VPN
:: Based off VPNUtilities
:: Returns error code 1 if jamCommand.exe is not found
:: See README.TXT for more info
:: mcarras8 7-10-25
SET "_TMPDIR=%SystemDrive%\TEMP"
SET "_AOLOCKFILE=Pulse Always On.lock"
SET "_DEFAULTLOCKFILE=Pulse Default.lock"
SET "_AOCONFIGFILE=Pulse Always On.pulsepreconfig"
SET "_DEFAULTCONFIGFILE=Default.pulsepreconfig"
If Not Exist "%_TMPDIR%" MD "%_TMPDIR%"
del "%_TMPDIR%\Config.pulsepreconfig"
:: If we're already on Always On VPN
If Exist "%_TMPDIR%\%_AOLOCKFILE%" (
	copy "%~dp0%_DEFAULTCONFIGFILE%" "%_TMPDIR%\Config.pulsepreconfig" /y
	del "%_TMPDIR%\%_AOLOCKFILE%"
	echo "Default VPN Config" > "%_TMPDIR%\%_DEFAULTLOCKFILE%"
) ELSE (
	copy "%~dp0%_AOCONFIGFILE%" "%_TMPDIR%\Config.pulsepreconfig" /y
	del "%_TMPDIR%\%_DEFAULTLOCKFILE%"
	echo "Always ON VPN Config" > "%_TMPDIR%\%_AOLOCKFILE%"
)

IF NOT EXIST "%ProgramFiles(x86)%\Common Files\Pulse Secure\JamUI\jamCommand.exe" (
	exit /b 1
) ELSE (
	pushd "%ProgramFiles(x86)%\Common Files\Pulse Secure\JamUI"
	jamCommand.exe /importfile "%_TMPDIR%\Config.pulsepreconfig"
	popd

	exit /b 0
)
