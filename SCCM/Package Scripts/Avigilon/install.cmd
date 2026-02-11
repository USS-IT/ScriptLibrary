@echo off
REM MJC 9-20-24
REM https://docs.avigilon.com/bundle/unity-video-software-manager-8-3/page/system-management/silent-install-command-options.htm
SET "_VmsInstallerExe=VmsInstaller-2024.6.1.1854-1720034519.exe"
SET "_InstallationManagerExe=%SystemDrive%\Program Files\Avigilon\Avigilon Unity Software Manager\UI\VmsInstallationManager.exe"
SET "_InstallationManagerExe32bit=%SystemDrive%\Program Files (x86)\Avigilon\Avigilon Unity Software Manager\UI\VmsInstallationManager.exe"
SET "_ClientExe="%SystemDrive%\Program Files\Avigilon\Avigilon Unity Client\VmsClientApp.exe"
pushd "%~dp0"
REM Install client software using Software Manager.
REM Uses a powershell script to handle cases where the Software Manager must be updated first.
echo [%~nx0] Installing Avigilon Unity Client Software...
powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "AvigilonVmsSilentInstallScript.ps1" -ExePath "%_InstallationManagerExe%" -CliArgs "install Client Player --acceptEula --sendUsageData=false --launchClient=false --silent" -VmsInstallerPath "%~dp0\%_VmsInstallerExe%"
REM Check if program got installed in a different path.
IF NOT EXIST "%_InstallationManagerExe%" (
	IF EXIST "%_InstallationManagerExe32bit%" (
		powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "AvigilonVmsSilentInstallScript.ps1" -ExePath "%_InstallationManagerExe32bit%" -CliArgs "install Client Player --acceptEula --sendUsageData=false --launchClient=false --silent" -VmsInstallerPath "%~dp0\%_VmsInstallerExe%"
	)
)
REM Check if the Client installed properly. If not, try again.
IF NOT EXIST "%_ClientExe%" (
	echo [%~nx0] Client not found. Reinstalling Avigilon Unity Client Software...
	powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "AvigilonVmsSilentInstallScript.ps1" -ExePath "%_InstallationManagerExe%" -CliArgs "install Client Player --acceptEula --sendUsageData=false --launchClient=false --silent" -VmsInstallerPath "%~dp0\%_VmsInstallerExe%"
)
popd
exit /b %ERRORLEVEL%
