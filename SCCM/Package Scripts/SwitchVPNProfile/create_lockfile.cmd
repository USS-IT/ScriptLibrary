@Echo off
:: Create the lock file for Always On VPN for first run.
:: This should only be run once.
:: mcarras8 7-28-25
SET "_TMPDIR=C:\temp"
SET "_AOLOCKFILE=Pulse Always On.lock"
If Not Exist "%_TMPDIR%" MD "%_TMPDIR%"
If Not Exist "%_TMPDIR%\%_AOLOCKFILE%" (
	echo "Always ON VPN Config" > "%_TMPDIR%\%_AOLOCKFILE%"
)
exit /b 0