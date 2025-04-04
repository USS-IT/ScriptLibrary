@echo off
REM This removes the old junction folders before the "DONOTREMOVE" tag was added.
REM --- Changelog ---
REM 03-20-25 - mcarras8 - Created.
REM ------------------

SET "_LINKDIR=C:\OUSA"
IF EXIST "%_LINKDIR%" (
	FOR /D %%G in ("%_LINKDIR%\*") DO (
		echo Removing old link [%%G]
		rmdir "%%~fG"
	)
	echo [%_LINKDIR%] has been cleaned. All set to run [USS_OUSA_SP_Junctions.bat].
	pause
) ELSE (
	echo [%_LINKDIR%] not found. All set to run [USS_OUSA_SP_Junctions.bat].
	pause
)