<#
.SYNOPSIS
    Sends a notification when it detects the MCM Client has successfully reinstalled.
.DESCRIPTION
    Sends a notification when it detects the MCM Client has successfully reinstalled, then unregisters its own task.
.OUTPUTS
	0 - Success
.NOTES
	Author: Matt Carras (mcarras8)
	Created: 07-28-2026
	
	Logs to C:\USS\Logs\User\MCMTroubleshooting\Watch-MCMClientReinstall.ps1.log by default.
	
	Requires Administrator privileges.
#>
#Requires -RunAsAdministrator
param(	
	[Switch] $Notify,

	[string] $LogDir = "C:\USS\Logs\User\MCMTroubleshooting"
)

# ----------------------------
# Configuration
# ----------------------------

# Amount of time to wait until the SCCM client finishes reinstalling.
$TimeoutMin = 30

# Name of the task that triggered this script.
$WatcherTaskName = "USS-Watch-MCMClientReinstall"

# -- END CONFIGURATION --

# Create the log path if it doesn't already exist
if (-not (Test-Path $LogDir -PathType Container)) {
	$null = New-Item -Path $LogDir -ItemType Directory
}
try {
	$_scriptName = Split-Path -Leaf $PSCommandPath
} catch {
	$_scriptName = "Watch-MCMClientReinstall.ps1"
}
$LogPath = "$LogDir\$($_scriptName).log"
Start-Transcript $LogPath -Force

# ----------------------------
# Helper functions
# ----------------------------

function Get-MCMVersion {
	# Return the current version of the MCM client, if it exists.
	param()
	
    $Paths = @(
        'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\SMS\Mobile Client'
    )

    foreach ($Path in $Paths) {
        try {
            $Version = (Get-ItemProperty `
                -Path $Path `
                -Name ProductVersion `
                -ErrorAction Stop).ProductVersion

            if (-not [string]::IsNullOrEmpty($Version)) {
                return $Version
            }
        }
        catch {}
    }

    return $null
}

function Get-SMSClient {
	# Return the SMS_Client class, if it exists.
	param()
	
    try {
		$owmi = Get-CimInstance -Namespace root\ccm `
						-Class SMS_Client `
						-ErrorAction Stop
						
		if ($owmi) {
			return $owmi
		}
	}
	catch {}

    return $null
}

# -- END FUNCTIONS --

$StartDateTime = (Get-Date)
$TimeoutDate = (Get-Date).AddMinutes($SetupTimeoutMin)

Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Watching for MCM client reinstall..."
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Criteria: CCMExec service must be running. MCM Client Product Version and SMS_Client WMI class must exist."
$smsClient = $null
$mcmVersion = $null
$svcCcmExec = $null
while ((Get-Date) -lt $TimeoutDate) {
	if ($mcmVersion -eq $null) {
		$mcmVersion = Get-MCMVersion
		if ($mcmVersion -ne $null) {
			Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] MCM Product Version exists. ProductVersion = $mcmVersion"
		}
	}
	
	if (-Not $smsClient) {
		$smsClient = Get-SMSClient
		if ($smsClient) {
			Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] SMS_Client WMI class exists."
		}
	}
	
	if ($svcCcmExe.Status -ne 'Running') {
		$svcCcmExe = Get-Service CcmExec -ErrorAction SilentlyContinue
		if ($svcCcmExe.Status -eq 'Running') {
			Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] CcmExec service exists and is currently running."
		}
	}
	
	if ($svcCcmExe.Status -eq 'Running' -And $mcmVersion -And $smsClient) {
		break
	}
	
	Start-Sleep 15
}

$reinstallSuccess = $false
$exitCode = 0
if ($svcCcmExe.Status -eq 'Running' -And $mcmVersion -And $smsClient) {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] MCM Client reinstalled successfully."
	$reinstallSuccess = $true
} elseif ((Get-Date) -lt $TimeoutDate) {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ERROR: Timed out waiting for MCM client reinstall."
	$exitCode = 1
} else {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ERROR: MCM Client may not have reinstalled successfully."
	$exitCode = 2
}

if ($Notify) {
	if ($reinstallSuccess) {
	}
}

# Unregister the scheduled task, if it exists.
if ((Get-ScheduledTask -TaskName $WatcherTaskName -ErrorAction SilentlyContinue)) {
	try {
		Unregister-ScheduledTask -TaskName $WatcherTaskName -Confirm:$false
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Unregistered watcher scheduled task."
	} catch {
		Write-Error $_
	}
}

Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Exiting with return code $exitCode"
Stop-Transcript | Out-Null
exit $exitCode