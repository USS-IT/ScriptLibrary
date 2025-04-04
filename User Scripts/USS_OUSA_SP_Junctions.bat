@echo off
REM Creates junction links for a universal path to a given directory or SharePoint site.
REM This can be run under a user account as junction links do not require special rights.
REM --- Changelog ---
REM 03-20-25 - mcarras8 - Added "_TARGETDIR" variable and made "_TARGETSITE" optional. Also added DONOTREMOVE tag when _TARGETSITE is used.
REM 10-11-23 - mcarras8 - Created
REM ------------------

REM Target directory. It will link this directory exactly.
REM Give either _TARGETSITE or _TARGETDIR, not both.
SET "_TARGETDIR=Office of Undergraduate Student Analytics - OUSA"
REM Target SharePoint site. It will search for all folders matching this pattern under the user's profile.
REM This will also create a "DONOTREMOVE" deployment file.
SET "_TARGETSITE="
SET "_LINKDIR=C:\OUSA"
SET "_TENANTNAME=Johns Hopkins"
IF NOT "%_TARGETSITE%" == "" (
	echo _TARGETSITE=%_TARGETSITE%
	mkdir "%_LINKDIR%" >Nul 2>&1
	FOR /D %%G in ("%_LINKDIR%\*") DO (
		echo Removing old link [%%G]
		rmdir "%%~fG"
	)
	FOR /D %%G in ("%USERPROFILE%\%_TENANTNAME%\%_TARGETSITE% -*") DO (
		mklink /J /D "%_LINKDIR%\%%~nG" "%%~fG"
	)
	echo TARGETSITE=%_TARGETSITE% > "%_LINKDIR%\DONOTREMOVE"
	echo Made by [%~n0] >> "%_LINKDIR%\DONOTREMOVE"
) ELSE (
	echo _TARGETDIR=%_TARGETDIR%
	IF NOT EXIST "%USERPROFILE%\%_TENANTNAME%\%_TARGETDIR%" (
		echo [%USERPROFILE%\%_TENANTNAME%\%_TARGETDIR%] not found. Aborting.
	) ELSE (
		REM If we have a DONOTREMOVE file, assume it was switched from using the _TARGETSITE parameter.
		IF EXIST "%_LINKDIR%\DONOTREMOVE" (
			FOR /D %%G in ("%_LINKDIR%\*") DO (
				echo Removing old link [%%G]
				rmdir "%%~fG"
			)
			del /f /q "%_LINKDIR%\DONOTREMOVE"
		)
		echo Removing old link [%_LINKDIR%]
		rmdir "%_LINKDIR%"
		mklink /J /D "%_LINKDIR%" "%USERPROFILE%\%_TENANTNAME%\%_TARGETDIR%"
	)
)
