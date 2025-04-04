@echo off
REM Copy files from one directory to another using robocopy.
REM Intended for copying from your HOME network drive to your OneDrive Documents folder.
REM mcarras8 4-4-25
SET "_SOURCEDIR=\\win.ad.jhu.edu\Users$\HOME"
SET "_DESTDIR=%OneDrive%\Documents"
IF NOT EXIST "%_DESTDIR%" (
	echo [%~nx0] ERROR: [%_DESTDIR%] does not exist. Exiting.
	GOTO :exit
) ELSE (
	IF NOT EXIST "%_SOURCEDIR%" (
		echo [%~nx0] ERROR: [%_SOURCEDIR%] does not exist. Please make sure you're connected to the VPN and run this script again.
		GOTO :exit
	) ELSE (
		echo [%~nx0] Copying HOME network drive files to [%_DESTDIR%], please wait...
		robocopy "\\win.ad.jhu.edu\Users$\HOME" "%_DESTDIR%" /E /Z /R:1 /W:5 /IPG:5 /XD "Windows" && echo ** All files have been copied.
	)
)
:exit
pause
