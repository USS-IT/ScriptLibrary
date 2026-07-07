@echo off
setlocal EnableDelayedExpansion
REM Creates the USS directory if it doesn't already exist with appropriate permissions.
REM These permissions should also be set by GPO.
REM 07-07-2026 Matt Carras (mcarras8)

REM Setup C:\USS folder structure.
SET "_MAINDIR=%SystemDrive%\USS"
SET "_USERSCRIPTSDIR=%_MAINDIR%\Scripts\User"
SET "_ADMINSCRIPTSDIR=%_MAINDIR%\Scripts\Admin"
SET "_LOGDIR=%_MAINDIR%\Logs"
SET "_USERLOGSDIR=%_LOGDIR%\User"
SET "_ADMINLOGSDIR=%_LOGDIR%\Admin"
SET "_LOGFILE=%_LOGDIR%\SetupUSSFolder.log"
REM Create and set permissions on main folder to read-only for users.
IF NOT EXIST "%_MAINDIR%" (
	echo Creating [%_MAINDIR%]...
	mkdir "%_MAINDIR%"
)
echo Setting permissions on [%_MAINDIR%]...
icacls "%_MAINDIR%" /inheritance:d /Q >Nul
icacls "%_MAINDIR%" /remove:g "Authenticated Users" /Q >Nul
icacls "%_MAINDIR%" /grant:r "Authenticated Users:(OI)(CI)(RX)" /Q >Nul
icacls "%_MAINDIR%" /remove:g "Users" /Q >Nul
icacls "%_MAINDIR%" /grant:r "Users:(OI)(CI)(RX)" /Q >Nul
REM Create and set permissions on main log folsder.
IF NOT EXIST "%_LOGDIR%" (
	echo Creating [%_LOGDIR%]...
	mkdir "%_LOGDIR%"
)
REM Log the rest of the output.
(
	echo ** Logging started %DATE% %TIME%
	
	IF NOT EXIST "%_USERLOGSDIR%" (
		echo Creating [%_USERLOGSDIR%]...
		mkdir "%_USERLOGSDIR%"
	) ELSE (
		echo [%_USERLOGSDIR%] already exists, not creating
	)
	REM The User logs folder needs to be writable by standard users.
	echo Setting permissions on [%_USERLOGSDIR%]...
	icacls "%_USERLOGSDIR%" /inheritance:d /Q >Nul
	icacls "%_USERLOGSDIR%" /remove:g "Authenticated Users" /Q >Nul
	icacls "%_USERLOGSDIR%" /grant:r "Authenticated Users:(M)" /Q >Nul
	icacls "%_USERLOGSDIR%" /grant "Authenticated Users:(OI)(CI)(IO)(F)" /Q >Nul
	icacls "%_USERLOGSDIR%" /remove:g "Users" /Q >Nul
	icacls "%_USERLOGSDIR%" /grant:r "Users:(M)" /Q >Nul
	icacls "%_USERLOGSDIR%" /grant "Users:(OI)(CI)(IO)(F)" /Q >Nul
	IF NOT EXIST "%_ADMINLOGSDIR%" (
		echo Creating [%_ADMINLOGSDIR%]...
		mkdir "%_ADMINLOGSDIR%"
	) ELSE (
		echo [%_ADMINLOGSDIR%] already exists, not creating
	)
	REM The Admin logs folder should only be readable by admins.
	echo Setting permissions on [%_ADMINLOGSDIR%]...
	icacls "%_ADMINLOGSDIR%" /inheritance:d /Q >Nul
	icacls "%_ADMINLOGSDIR%" /remove:g "Authenticated Users" /Q >Nul
	icacls "%_ADMINLOGSDIR%" /remove:g "Users" /Q >Nul
	
	REM Create user scripts folder.
	IF NOT EXIST "%_USERSCRIPTSDIR%" (
		echo Creating [%_USERSCRIPTSDIR%]...
		mkdir "%_USERSCRIPTSDIR%"
	) ELSE (
		echo [%_USERSCRIPTSDIR%] already exists, not creating
	)
	REM Create admin scripts folder and set permissions.
	IF NOT EXIST "%_ADMINSCRIPTSDIR%" (
		echo Creating [%_ADMINSCRIPTSDIR%]...
		mkdir "%_ADMINSCRIPTSDIR%"
	) ELSE (
		echo [%_ADMINSCRIPTSDIR%] already exists, not creating
	)
	REM The Admin scripts folder should only be readable by admins.
	echo Setting permissions on [%_ADMINSCRIPTSDIR%]...
	icacls "%_ADMINSCRIPTSDIR%" /inheritance:d /Q >Nul
	icacls "%_ADMINSCRIPTSDIR%" /remove:g "Authenticated Users" /Q >Nul
	icacls "%_ADMINSCRIPTSDIR%" /remove:g "Users" /Q >Nul
	
	echo ERRORLEVEL=!ERRORLEVEL!
) > "%_LOGFILE%" 2>&1
EXIT /b !ERRORLEVEL!