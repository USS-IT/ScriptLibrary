@echo off
REM Disables Spotlight from being the default background for new users.
REM See https://learn.microsoft.com/en-us/windows/configuration/windows-spotlight/?pivots=windows-11 note about KB5046633 changes
REM Adapted from source: https://www.reddit.com/r/Intune/comments/1hfvjou/setting_picture_to_the_default_for_new_users/
REM mcarras8 1-10-25

REM Load the NTUSER.DAT into HKU\TEMP
REG LOAD HKU\Temp "%SystemDrive%\users\default\ntuser.dat"
  
REM Create new registry key...  
Reg Add "HKU\TEMP\Software\Microsoft\Windows\CurrentVersion\DesktopSpotlight\Settings" /f

REM Prevent this change from happening...
Reg Add "HKU\TEMP\Software\Microsoft\Windows\CurrentVersion\DesktopSpotlight\Settings" /f /v "OneTimeUpgrade" /t Reg_DWORD /d 1

REM Reverting back to Picture as desktop customization type...
Reg Add "HKU\TEMP\Software\Microsoft\Windows\CurrentVersion\DesktopSpotlight\Settings" /f /v "EnabledState" /t Reg_DWORD /d 0

REM Remove Temp hive and unload the registry hive HKU\TEMP
REG UNLOAD HKU\TEMP

exit /b

