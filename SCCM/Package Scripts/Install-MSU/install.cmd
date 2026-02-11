@echo off
REM 2025-10 Cumulative Update for .NET Framework 3.5 and 4.8.1 for Windows 11, version 23H2 for x64 (KB5066133)
REM Author: Matt Carras (mcarras8)
REM Created: 2-6-26
REM powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Install-MSU.ps1" -MSU "%~dp0windows11.0-kb5066133-x64-ndp481_4dd0fbfc976c5ef5b9492ab42895e8b56310f430.msu" -InstalledPackage "Package_for_DotNetRollup_481~31bf3856ad364e35~amd64~~10.0.9320.2"
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Install-MSU.ps1" -MSU "%~dp0windows11.0-kb5066133-x64-ndp481_4dd0fbfc976c5ef5b9492ab42895e8b56310f430.msu"
exit /b %ERRORLEVEL%