@echo off
setlocal EnableDelayedExpansion
REM Creates the USS directory if it doesn't already exist with appropriate permissions.
REM These permissions should also be set by GPO.
REM It then copies over the scripts and installs the scheduled tasks.
REM 07-28-2026 Matthew Carras (mcarras8)

REM Setup C:\USS folder structure.
SET "_MAINDIR=%SystemDrive%\USS"
SET "_USERSCRIPTSDIR=%_MAINDIR%\Scripts\User"
SET "_ADMINSCRIPTSDIR=%_MAINDIR%\Scripts\Admin"
SET "_LOGDIR=%_MAINDIR%\Logs"
SET "_USERLOGSDIR=%_LOGDIR%\User"
SET "_ADMINLOGSDIR=%_LOGDIR%\Admin"
SET "_STARTMENUFOLDER=%ProgramData%\Microsoft\Windows\Start Menu\Programs\USS"
SET "_PACKAGEFOLDER=MCMTroubleshooting"
SET "_LOGFILE=%_LOGDIR%\Packages\%_PACKAGEFOLDER%\Install_Package.log"
REM Set permissions on main folder to read-only for users.
IF NOT EXIST "%_MAINDIR%" (
	echo Creating [%_MAINDIR%]...
	mkdir "%_MAINDIR%"
	icacls "%_MAINDIR%" /inheritance:d /Q >Nul
	icacls "%_MAINDIR%" /remove:g "Authenticated Users" /Q >Nul
	icacls "%_MAINDIR%" /grant:r "Authenticated Users:(OI)(CI)(RX)" /Q >Nul
	icacls "%_MAINDIR%" /remove:g "Users" /Q >Nul
	icacls "%_MAINDIR%" /grant:r "Users:(OI)(CI)(RX)" /Q >Nul
)
REM Set permissions on logs folders.
IF NOT EXIST "%_USERLOGSDIR%" (
	echo Creating [%_USERLOGSDIR%]...
	mkdir "%_USERLOGSDIR%"
	icacls "%_USERLOGSDIR%" /inheritance:d /Q >Nul
	icacls "%_USERLOGSDIR%" /remove:g "Authenticated Users" /Q >Nul
	icacls "%_USERLOGSDIR%" /grant:r "Authenticated Users:(M)" /Q >Nul
	icacls "%_USERLOGSDIR%" /grant "Authenticated Users:(OI)(CI)(IO)(F)" /Q >Nul
	icacls "%_USERLOGSDIR%" /remove:g "Users" /Q >Nul
	icacls "%_USERLOGSDIR%" /grant:r "Users:(M)" /Q >Nul
	icacls "%_USERLOGSDIR%" /grant "Users:(OI)(CI)(IO)(F)" /Q >Nul
)
mkdir "%_USERLOGSDIR%\%_PACKAGEFOLDER%"
IF NOT EXIST "%_ADMINLOGSDIR%" (
	echo Creating [%_ADMINLOGSDIR%]...
	mkdir "%_ADMINLOGSDIR%"
	icacls "%_ADMINLOGSDIR%" /inheritance:d /Q >Nul
	icacls "%_ADMINLOGSDIR%" /remove:g "Authenticated Users" /Q >Nul
	icacls "%_ADMINLOGSDIR%" /remove:g "Users" /Q >Nul
)
mkdir "%_MAINDIR%\Logs\Packages\%_PACKAGEFOLDER%"
REM Log the rest of the output.
(
	echo ** Started %DATE% %TIME%
	REM Create user scripts folder.
	echo Creating [%_USERSCRIPTSDIR%\%_PACKAGEFOLDER%]...
	mkdir "%_USERSCRIPTSDIR%\%_PACKAGEFOLDER%"
	REM Set permissions on admin scripts folder.
	IF NOT EXIST "%_ADMINSCRIPTSDIR%" (
		echo Creating [%_ADMINSCRIPTSDIR%]...
		mkdir "%_ADMINSCRIPTSDIR%"
		icacls "%_ADMINSCRIPTSDIR%" /inheritance:d /Q >Nul
		icacls "%_ADMINSCRIPTSDIR%" /remove:g "Authenticated Users" /Q >Nul
		icacls "%_ADMINSCRIPTSDIR%" /remove:g "Users" /Q >Nul
	)
	REM Create the folder for this package.
	echo Creating [%_ADMINSCRIPTSDIR%\%_PACKAGEFOLDER%]...
	mkdir "%_ADMINSCRIPTSDIR%\%_PACKAGEFOLDER%"
	REM Create the start menu folder for any shortcuts.
	echo Creating [%_STARTMENUFOLDER%]...
	mkdir "%_STARTMENUFOLDER%"

	REM Copy files and install scheduled tasks.
	echo Copying [Admin\*] to [%_ADMINSCRIPTSDIR%\%_PACKAGEFOLDER%]...
	copy "%~dp0Admin\*" "%_ADMINSCRIPTSDIR%\%_PACKAGEFOLDER%" /V /Y
	echo Copying [User\*] to [%_USERSCRIPTSDIR%\%_PACKAGEFOLDER%]...
	copy "%~dp0User\*" "%_USERSCRIPTSDIR%\%_PACKAGEFOLDER%" /V /Y
	
	REM Install Scheduled Tasks
	powershell.exe -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0Install-Tasks.ps1"
	IF !ERRORLEVEL! NEQ 0 (
		echo FATAL ERROR CODE [!ERRORLEVEL!], aborting
	) ELSE (
		echo Copying [Shortcuts\*] to [%_STARTMENUFOLDER%]...
		copy "%~dp0Shortcuts\*" "%_STARTMENUFOLDER%" /V /Y
	)
	
	echo ERRORLEVEL=!ERRORLEVEL!
) > "%_LOGFILE%" 2>&1
EXIT /b !ERRORLEVEL!