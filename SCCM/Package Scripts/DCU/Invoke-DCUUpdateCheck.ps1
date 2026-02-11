<#
	.SYNOPSIS
	Script to trigger Dell Command Update manually.
	
	.DESCRIPTION
	Script to trigger Dell Command Update manually to install bios/firmware/driver/other updates delivered by Dell.
	
	.PARAMETER UpdateType
	Optional update type to install from DCU in a single, comma-delimited string. If not given, installs all updates DCU is set to allow by default. Valid types: bios,firmware,driver,application,others

	.PARAMETER DriverPack
	Install the given driver pack instead of scanning for available updates from Dell's servers.
		
	.PARAMETER CheckInstallAndSleep
	Waits a preset amount of time if the logfile indicates any updates were interrupted applying that day. Can be used after rebooting to make sure DCU is finished applying all updates.
	
	.PARAMETER RerunIfPreviousFailure
	Reruns DCU with given parameters if the logfile indicates any updates failed to apply that day.
	
	.PARAMETER ShowGUI
	Disables the silent option.
	
	.PARAMETER LogDir
	Directory to store / look for log files. Will create directory if not found. Defaults to C:\Dell\Logs
	
	.PARAMETER ReturnExitCode
	Always return the exit code from DCU.
	
	.PARAMETER ReturnExitCodeOnFailure
	Return the exit code from DCU on failure instead of always exiting with 3010.
	
	.OUTPUTS
	Exit code 0 - No action taken
	Exit code 1 - Cannot find DCU installation
	Exit code 2 - Not a Dell system
	Exit code 1618 - AC power is not detected
	Exit code 3010 - DCU successful, restart required
	
	.EXAMPLE
	Invoke-DCUUpdateCheck.ps1 -UpdateType "bios,firmware" -RerunIfPreviousFailure
	
	.NOTES
	Author: mcarras8 2-2-24
	
	DCU should be installed with update check set to Manual and all options for BIOS,Firmware, and Drivers enabled.
	
	You can check for exit codes using _SMSTSLastActionRetCode in a task sequence.
#>
param( 
	[Parameter(Mandatory=$false)]
	[ValidateScript({$_ -eq $null -Or (($_ -split ",") | where {$_ -notin @("bios","firmware","driver","application","others")}).Count -le 0})]
	[string]$UpdateType,
	
	[string]$DriverPack,
		
	[switch]$CheckInstallAndSleep,
	
	[switch]$RerunIfPreviousFailure,

	[switch]$ShowGUI,
	
	[string]$LogDir="C:\Dell\Logs",
	
	[switch]$ReturnExitCode,
	
	[switch]$ReturnExitCodeOnFailure
)


# Directory for logs.
# Saves to C:\Dell\Logs by default.
# Note: C:\Windows paths cannot be used due to that being a "reserved" path.
# If local path for logs doesn't exist, create it
If (!(Test-Path $LogDir)) { 
	New-Item -Path $LogDir -Type Directory -Force 
}
$logFileName = ("DCU_{0}.log" -f (Get-Date -Format "yyyy_MM_dd"))
$logFilePath = "$LogDir\$logFileName"
 
#Check for AC power and exit if missing
Add-Type -Assembly System.Windows.Forms
$PowerStatus = [System.Windows.Forms.SystemInformation]::PowerStatus
If ($PowerStatus.PowerLineStatus -eq "Offline") {
	Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] ERROR: PowerLineStatus = Offline." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss")) -PassThru
	exit 1618
}
 
#Make sure device is actually a dell.
$PCInfo = Get-WMIObject -Query "Select * from Win32_ComputerSystem" | Select-Object -Property Manufacturer, Model

#Execute Dell Command Update
$finalExitCode = 0
if ($PCInfo.Manufacturer -like "Dell*" -OR $PCInfo.Manufacturer -like "Alienware*"){
	$doRun = $true
	If($RerunIfPreviousFailure) {
		# Only rerun if previous failure is detected.
		$doRun = $false
		If((Test-Path $logFilePath -PathType Leaf)) {
			$logContent = Get-Content "$logFilePath"
			If ($logContent -And $logContent -match ("\[{0}.+ update\(s\) failed to install" -f (Get-Date -Format "yyyy-MM-dd"))) {
				$doRun = $true
			}
		}
	}
	
	# Checks to see if we need to wait for installation of some updates to finish before uninstalling DCU.
	If($CheckInstallAndSleep) {
		$sleepValue = 300
		# Try to read the log file for clues on how the installation went.
		If((Test-Path $logFilePath -PathType Leaf)) {
			$logContent = Get-Content "$logFilePath"
			If ($logContent) {
				If ($logContent -match ("\[{0}.+ Installation for these updates were interrupted" -f (Get-Date -Format "yyyy-MM-dd"))) {
					# Wait longer if installation was interrupted.
					$sleepValue = 830
				<#
				} elseif ($logContent -match ("\[{0}.+ The system has been updated" -f (Get-Date -Format "yyyy-MM-dd"))) {
					# Do not wait if all updates were successful.
					$sleepValue = $null
				#>
				}
			}
		}
		if ($sleepValue) {
			Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] Sleeping for {1} seconds while waiting for installs to finish..." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $sleepValue) -PassThru
			Start-Sleep $sleepValue
		}
	} elseif($doRun) {		
		$DCUexe = "${ENV:ProgramFiles}\Dell\CommandUpdate\dcu-cli.exe"
		If (!(Test-Path $DCUexe)) {
			$DCUexe = "${ENV:ProgramFiles(x86)}\Dell\CommandUpdate\dcu-cli.exe"
			If (!(Test-Path $DCUexe)) {
				Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] ERROR: Dell Command Update is not installed." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss")) -PassThru
				exit 1
			}
		}
		
		# Attempt to install the given Driver Pack.
		if (-Not [string]::IsNullOrEmpty($DriverPack)) {
			# Set the configuration to allow advanced driver restore
			$DCUparameters = "/configure -advancedDriverRestore=enable -outputLog=$logFilePath"
			$Params = $DCUparameters.Split(" ")
			Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] $DCUexe $Params" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"))
			$output = & $DCUexe $Params
			Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] Dell Command Update returned exit code {1}" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $LASTEXITCODE) -PassThru
			Add-Content -Path $logFilePath -Value $output -PassThru
			
			# Attempt to load the given driver pack.
			$DCUparameters = "/driverInstall -driverLibraryLocation=""$DriverPack"" -outputLog=$logFilePath -reboot=disable"
			If(-Not $ShowGUI) {
				$DCUparameters += " -silent"
			}
		} else {
			# Check for updates from Dell.
			
			# Run a scan first.
			$DCUparameters = "/scan -outputLog=$logFilePath"
			If(-Not $ShowGUI) {
				$DCUparameters += " -silent"
			}
			$Params = $DCUparameters.Split(" ")
			Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] $DCUexe $Params" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"))
			$output = & $DCUexe $Params
			$dcuExitCode = $LASTEXITCODE
			Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] Dell Command Update returned exit code {1}" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $dcuExitCode) -PassThru
			Add-Content -Path $logFilePath -Value $output -PassThru
			
			$DCUparameters = "/applyUpdates -AutoSuspendBitlocker=enable -outputLog=$logFilePath -reboot=disable -forceupdate=enable"
			If(-Not $ShowGUI) {
				$DCUparameters += " -silent"
			}
			
			If($UpdateType) {
				$DCUparameters += " -updateType=$UpdateType"
			}
		}
		
		$Params = $DCUparameters.Split(" ")
		Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] $DCUexe $Params" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"))
		$output = & $DCUexe $Params
		$dcuExitCode = $LASTEXITCODE
		Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] Dell Command Update returned exit code {1}" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $dcuExitCode) -PassThru
		Add-Content -Path $logFilePath -Value $output -PassThru
		if ($ReturnExitCode -Or ($ReturnExitCodeOnFailure -And $dcuExitCode -lt 0)) {
			$finalExitCode = $dcuExitCode
		} else {
			$finalExitCode = 3010
		}
	}
} else {
	Add-Content -Path $logFilePath -Value ("[{0}] [Invoke-DCUUpdateCheck.ps1] Does not appear to be a Dell system." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss")) -PassThru
	$finalExitCode = 2
}

exit $finalExitCode
