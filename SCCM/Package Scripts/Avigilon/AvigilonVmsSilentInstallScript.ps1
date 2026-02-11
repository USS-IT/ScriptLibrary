<#
.SYNOPSIS
    Execute a Software Manager CLI command with auto-update handling
.PARAMETER ExePath
    The full path to the Avigilon Unity Software Manager installer or application executable
.PARAMETER CliArgs
    CLI arguments to pass
.PARAMETER TimeoutSeconds
    Maximum time to allow the Avigilon Unity Software Manager to run the command before aborting
.PARAMETER VmsInstallerPath
	The full path to the installer for the Aviligon Unity Software Manager software
.PARAMETER LogFilename
	The filename for the log file, default: AvigilonVmsSilentInstallScript.ps1.log
.PARAMETER LogDir
	The directory path to save the log file, default: C:\Temp\Logs
.NOTES
    Copyright (c) Motorola Solutions Incorporated. All Rights Reserved.

    Unauthorized reproduction, use or disclosure, in whole or in part, of this file and its contents, without
    express written consent of Motorola Solutions Inc., is prohibited.
#>
#Requires -RunAsAdministrator
param (
    [Parameter(Mandatory)]
    [string]$ExePath,
    [Parameter(Mandatory)]
    [string[]]$CliArgs,
    [int]$TimeoutSeconds = 1800,
	[string]$VmsInstallerPath,
	[string]$LogFilename="AvigilonVmsSilentInstallScript.ps1.log",
	[string]$LogDir="C:\TEMP\Logs"
)

If (-not [string]::IsNullOrWhitespace($LogFilename)) {
	If (!(Test-Path $LogDir)) { 
		New-Item -Path $LogDir -Type Directory -Force 
	}
	Start-Transcript -Path "$LogDir\$LogFilename" -Append -Force
}

# Install Vms if not already installed.
if (-not ([System.IO.File]::Exists($ExePath))) {
	$throwError = $true
	If (-not [string]::IsNullOrWhitespace($VmsInstallerPath)) {
		If (-not ([System.IO.File]::Exists($VmsInstallerPath))) {
			Write-Warning "Could not locate Vms Installer at ""$VmsInstallerPath"""
		} else {
			Write-Host "Located Vms Installer ""$($VmsInstallerPath)""`nRunning executable with CLI Args ""install,--silent"""
			$process = Start-Process -FilePath "$VmsInstallerPath" -ArgumentList "install","--silent" -NoNewWindow -PassThru
			$process | Wait-Process -Timeout $TimeoutSeconds
			$counter = 0
			$sleepTime = 10
			$maxSleepTime = 300
			if (-not ([System.IO.File]::Exists($ExePath))) {
				Write-Host "Installer still not found at ""$ExePath"". Waiting up to $maxSleepTime seconds for installer to finish..."
			}
			while ($counter -le $maxSleepTime -And -not ([System.IO.File]::Exists($ExePath))) {
				Start-Sleep $sleepTime
				$counter += $sleepTime
			}
			if (([System.IO.File]::Exists($ExePath))) {
				$throwError = $false
			}
		}
	}
	if ($throwError) {
		Write-Error "Failed to locate executable ""$ExePath"""
		exit 1
	}
}

# Keep trying up to the max retry count.
$maxRetryCount = 3
$retryCount = 0
while($retryCount -lt $maxRetryCount) {
	Write-Host "Located executable ""$($ExePath)""`nRunning executable with CLI Args ""$CliArgs"""
	$process = Start-Process -FilePath "$ExePath" -ArgumentList $CliArgs -NoNewWindow -PassThru
	# Saving the handle will allow us to query the exit code later.
	$handle = $process.Handle
	$process | Wait-Process -Timeout $TimeoutSeconds
	if ($process.ExitCode -eq 350 -or $process.ExitCode -eq 183) {
		Write-Host "Avigilon Unity Software Manager requires an update or is already running (exitcode: $($process.ExitCode)), searching for the update process for up to 5 minutes"
		$sleepCounter = 0
		$process = Get-Process | Where {$_.ProcessName -Like "Avigilon Unity Software Manager*"}
		$handle = $process.Handle
		while ($sleepCounter -lt 300 -And -Not $process) {
			Start-Sleep 10
			$sleepCounter += 10
			$process = Get-Process | Where {$_.ProcessName -Like "Avigilon Unity Software Manager*"}
			$handle = $process.Handle
		}
		if ($process) {
			Write-Host "Found update process ""$($process.ProcessName)""`nWaiting for update to complete"
			$process | Wait-Process -Timeout $TimeoutSeconds
			if ($process.ExitCode -eq 0) {
				Write-Host "Avigilon Unity Software Manager update was successful`nRunning the executable again with CLI Args ""$CliArgs"""
				$process = Start-Process -FilePath "$ExePath" -ArgumentList $CliArgs -NoNewWindow -PassThru
				$handle = $process.Handle
				$process | Wait-Process -Timeout $TimeoutSeconds
				# Break out early
				$retryCount = $maxRetryCount
			} else {
				$retryCount++
				Write-Error "Avigilon Unity Software Manager returned exit code [$($process.ExitCode)]! Retrying $retryCount out of $maxRetryCount."
				Start-Sleep 60
			}
		} else {
			$retryCount++
			Write-Error "Avigilon Unity Software Manager update process could not be found! Retrying $retryCount out of $maxRetryCount."
			Start-Sleep 60
		}
	}
}

if ($process.ExitCode -ne 0) {
    Write-Host "Avigilon Unity Software Manager exited with code: $($process.ExitCode)"    
}
else {
    Write-Host "Avigilon Unity Software Manager exited successfully"
}

Stop-Transcript

exit $process.ExitCode

