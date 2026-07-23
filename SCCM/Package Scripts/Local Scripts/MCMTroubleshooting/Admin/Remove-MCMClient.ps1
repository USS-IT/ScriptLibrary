<#
.SYNOPSIS
    Uninstalls the SCCM client, then optionally restarts.
.DESCRIPTION
    Uninstalls the SCCM client, then optionally restarts.
.PARAMETER Notify
	Invokes an event-based notification script to notify the current user after restart.
.PARAMETER NoRestart
	Exit instead of restarting.
.PARAMETER LogDir
	Directory for logging. Default is C:\USS\Logs.
.OUTPUTS
	0 - Success
	>0 - Error codes from ccmsetup.exe
	-1 - Error starting ccmsetup.exe
	-2 - Ccmexec service still exists after uninstall finished
.NOTES
	Author: Matt Carras (mcarras8)
	Created: 07-23-2026
	
	Logs to C:\USS\Logs\Remove-MCMClient.ps1.log by default.
	
	Requires Administrator privileges.
#>
#Requires -RunAsAdministrator
param(
	[switch] $Notify,
	
	[switch] $NoRestart,
	
	[string] $LogDir = "C:\USS\Logs"
)

$SetupTimeoutMin = 30
$RestartSec = 300
$WmiTimeoutMin = 5
$CCMLogPath = "$env:windir\ccmsetup\Logs\ccmsetup.log"

# Create the log path if it doesn't already exist
if (-not (Test-Path $LogDir -PathType Container)) {
	$null = New-Item -Path $LogDir -ItemType Directory
}
try {
	$_scriptName = Split-Path -Leaf $PSCommandPath
} catch {
	$_scriptName = "Remove-MCMClient.ps1"
}
$LogPath = "$LogDir\$($_scriptName).log"
Start-Transcript $LogPath -Force

function Get-CCMSetupReturnCode {
	# Return the return code from ccmsetup if it exited after the recorded start time, otherwise return $false.
    param(
		[Parameter(Mandatory, Position=0)]
        [string]$LogPath,
		[Parameter(Mandatory, Position=1)]
		[Alias('Start')]
        [datetime]$StartDateTime
    )

    if (-not (Test-Path $LogPath)) {
        return $false
    }

    $Line = Select-String -Path $LogPath -SimpleMatch 'CcmSetup is exiting with return code' | where {$_ -match 'return code (\d+)\].+<time="(\d{1,2}:\d{1,2}:\d{1,2}(?:\.\d+)?)(?:\+\d+)?" date="([^"]+)"'} | Select -Last 1
	if ($Matches) {
        try {
			$returnCode = $Matches[1] -as [int]
			$StrDateTime = "$($Matches[3]) $($Matches[2])"
            $LogDateTime = [datetime]::ParseExact(
                $StrDateTime,
                "MM-dd-yyyy HH:mm:ss.fff",
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
        catch {
            throw $_
        }

        if ($LogDateTime -ge $StartDateTime) {
            return $returnCode
        }
    }

    return $false
}

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

Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Polling initial state"
$mcmVersion = Get-MCMVersion
$oSmsClient = Get-SMSClient
$svcCcmExec = Get-Service CcmExec -ErrorAction SilentlyContinue
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] MCM Client Version (Registry): $mcmVersion"
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Ccmexec service status: $($svcCcmExec.Status)"
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] SMS_Client WMI Class Version: $($oSmsClient.ClientVersion)"

Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Calling ccmsetup.exe /uninstall"

$StartDateTime = (Get-Date)
$TimeoutDate = (Get-Date).AddMinutes($SetupTimeoutMin)
try {
	$process = Start-Process `
		-FilePath "$env:windir\ccmsetup\ccmsetup.exe" `
		-ArgumentList "/uninstall" `
		-WindowStyle Hidden `
		-ErrorAction Stop
	$handle = $process.Handle
	$process | Wait-Process -Timeout ($SetupTimeoutMin*60)
	$ccmExitCode = $process.ExitCode
	if ($ccmExitCode -is [int]) {
		$ccmExitCode = "exit code $ccmExitCode"
	} else {
		$ccmExitCode = "no exit code"
	}
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ccmsetup.exe /uninstall exited with $ccmExitCode. Uninstall should be continuing in the background. Exiting with code -1."
} catch {
	Write-Error $_
	Stop-Transcript | Out-Null
	exit -1
}

Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Waiting for MCM client uninstall to complete..."

$logPathExists = $false
while ((Get-Date) -lt $TimeoutDate) {
    if ((Test-Path $CCMLogPath)) {
		if (-Not $logPathExists) {
			$logPathExists = $true
			Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Log path [$CCMLogPath] exists. Watching..."
		}
		$returnCode = Get-CCMSetupReturnCode $CCMLogPath $StartDateTime
		
        if ($returnCode -is [int]) {
			if ($returnCode -eq 0) {
				Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Ccmsetup exited with return code 0. MCM client uninstall completed successfully."
				break
			} elseif ($returnCode -eq 8) {
				Write-Error "MCM client uninstall unsuccessful. Ccmsetup exited with return code $returnCode (ccmsetup already running). Check ccmsetup.log for more info. Returning ccmsetup exit code."
				Stop-Transcript | Out-Null
				exit $returnCode
			} else {
				Write-Error "MCM client uninstall unsuccessful. Ccmsetup exited with return code $returnCode. Check ccmsetup.log for more info. Returning ccmsetup exit code."
				Stop-Transcript | Out-Null
				exit $returnCode
			}
		}
    }

    Start-Sleep -Seconds 15
}

if ((Get-Date) -ge $TimeoutDate) {
    Write-Warning "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] MCM client uninstall wait reached $SetupTimeoutMin minute timeout."
}

Start-Sleep -Seconds 30

# Additional validation before restart
# Wait up to $WmiTimeoutMin (5 minutes) for WMI class to be removed.
# This also helps wait for additional cleanup.
$smsClientRemoved = $false
$WmiTimeoutDate = (Get-Date).AddMinutes($WmiTimeoutMin)
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Waiting for SMS_Client WMI class to be removed..."
while ((Get-Date) -lt $WmiTimeoutDate) {
	$oSmsClient = Get-SMSClient
	if (-Not $oSmsClient) {
		$smsClientRemoved = $true
		break
	}
	Start-Sleep -Seconds 15
}
if ($smsClientRemoved) {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Confirmed SMS_Client WMI class was successfully removed by ccmsetup."
} else {
	Write-Warning "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] SMS_Client WMI class still exists with version [$($oSmsClient.ClientVersion)] after MCM client uninstall attempt. Waited $WmiTimeoutMin minutes."
}

$svcCcmExec = Get-Service CcmExec -ErrorAction SilentlyContinue
if (-not $svcCcmExec) {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Confirmed ccmexec service was successfully removed by ccmsetup."
} else {
    Write-Warning "CcmExec service still exists after MCM client uninstall attempt (current status: $($svcCcmExec.Status). This may prevent the reinstall script from triggering."
}

$mcmVersion = Get-MCMVersion
if ([string]::IsNullOrEmpty($mcmVersion)) {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] MCM Client Version registry entry was successfully wiped by ccmsetup."
} else {
	Write-Warning "MCM client version is still set as [$mcmVersion] in registry after uninstall attempt. This may prevent the reinstall script from triggering."
}

if ($svcCcmExec -And [string]::IsNullOrEmpty($mcmVersion)) {
	Write-Error "ccmexec service and MCM client version both still exist after uninstall. This will prevent the MCM client script from running on startup. Try running this script again. Exiting with return code -2."
	Stop-Transcript | Out-Null
	exit -2
}

if ($NoRestart) {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] -NoRestart given. Exiting with return code 0."
} else {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Restarting in $($RestartSec / 60) minutes."
	shutdown.exe /r /t $RestartSec /c "SCCM client removal completed. Please restart the system. If the system is not restarted sooner, it will automatically restart in $($RestartSec / 60) minutes."
}

Stop-Transcript | Out-Null
exit 0