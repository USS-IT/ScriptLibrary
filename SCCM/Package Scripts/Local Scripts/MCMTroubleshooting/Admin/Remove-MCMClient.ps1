<#
.SYNOPSIS
    Uninstalls the SCCM client, then restarts.
.DESCRIPTION
    Uninstalls the SCCM client, then restarts.
	
	- Watches the ccmsetup.log file to determine when the client is done uninstalling. Error if both the service and registry product version still exist after an uninstall attempt.
	- The main reinstall script checks if either Ccmexec does not exist, or the MCM client version in the registry is older or does not exist.
	- If the script can't find ccmsetup.exe to start the process, but ccmexec and the client version still exist, it will attempt to remove the ccmexec service.
.PARAMETER NoRestart
	Exit instead of restarting.
.PARAMETER CodeFile
	Check the given code file path against the weekly code (from Get-WeeklyRemovalCode.ps1). The process will only start if the file has a valid code.
.PARAMETER LogDir
	Directory for logging. Default is C:\USS\Logs.
.OUTPUTS
	0 - Success
	>0 - Failure codes from ccmsetup.exe
	-1 - Failed to remove ccmexec service after confirming ccmsetup.exe does not exist
	-2 - Ccmexec service and MCM client version still exist after uninstall finished
	-3 - Invalid code file.
.NOTES
	Author: Matt Carras (mcarras8)
	Created: 07-23-2026
	
	Logs to C:\USS\Logs\Remove-MCMClient.ps1.log by default.
	
	If a CodeFile is given, the process will only start if the code file exists and is valid.
	
	Requires Administrator privileges.
#>
#Requires -RunAsAdministrator
param(	
	[switch] $NoRestart,
	
	[string] $CodeFile,
	
	[string] $LogDir = "C:\USS\Logs"
)

# ----------------------------
# Configuration
# ----------------------------

# Amount of time to watch the ccmsetup logs for a valid return code.
$SetupTimeoutMin = 30
# Amount of time before computer automatically restarts after determining the client was successfully uninstalled.
$RestartSec = 300
# Amount of time to additionally wait for the SMS_Client WMI class to be removed.
$WmiTimeoutMin = 5
# Path to Windows's ccmsetup.log and ccmsetup.exe files.
$CCMLogPath = "$env:windir\ccmsetup\Logs\ccmsetup.log"
$CCMExePath = "$env:windir\ccmsetup\ccmsetup.exe"

$WatcherTask
# -- END CONFIGURATION --

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

# ----------------------------
# Helper functions
# ----------------------------

function Get-WeeklyCode {
	# Return a 4 digit code that rotates every Sunday.
	param()
	
	$Secret = 'USS-MCM-Removal-Salt'

    $Today = (Get-Date).ToUniversalTime()

    # Get this week's Sunday
    $WeekStart = $Today.Date.AddDays(-[int]$Today.DayOfWeek)

    # Example: 20260719
    $WeekKey = $WeekStart.ToString('yyyyMMdd')

	# Get computer name, all uppercase
	$compname = ($env:COMPUTERNAME).Trim().ToUpper()
	
    # Combine with secret
    $InputString = "$WeekKey|$Secret|$compname"

    # SHA256 hash
    $Sha = [System.Security.Cryptography.SHA256]::Create()
    $Bytes = [System.Text.Encoding]::UTF8.GetBytes($InputString)
    $Hash = $Sha.ComputeHash($Bytes)

    # Convert first 4 bytes to an integer
	$Number = [BitConverter]::ToUInt32($Hash, 0)

    # Force into range 0000-9999
    $Code = $Number % 10000

    return $Code.ToString('0000')
}

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

# -- END FUNCTIONS --

# Check if we require a code file to continue.
if (-Not [string]::IsNullOrEmpty($CodeFile)) {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Code file given. Checking code."
	$validCode = $false
	if ((Test-Path $CodeFile)) {
		$code = Get-Content $CodeFile -Encoding ASCII
		$validCode = (-Not [string]::IsNullOrEmpty($code) -And $code -eq (Get-WeeklyCode))		
	}
	if ($validCode) {
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Code is valid. Continuing."
	} else {
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] FATAL ERROR: Invalid code. Aborting."
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Exiting with return code -3"
		Stop-Transcript | Out-Null
		exit -3
	}
}
	
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Polling initial state"
$mcmVersion = Get-MCMVersion
$oSmsClient = Get-SMSClient
$svcCcmExec = Get-Service CcmExec -ErrorAction SilentlyContinue
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] MCM Client Version (Registry): $mcmVersion"
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Ccmexec service status: $($svcCcmExec.Status)"
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] SMS_Client WMI Class Version: $($oSmsClient.ClientVersion)"

$StartDateTime = (Get-Date)
$TimeoutDate = (Get-Date).AddMinutes($SetupTimeoutMin)

Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Calling ccmsetup.exe /uninstall"
$ccmSetupError = $false
$ccmSetupMissing = $false
if ((Test-Path $CCMExePath)) {
	try {
		$process = Start-Process `
			-FilePath $CCMExePath `
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
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ccmsetup.exe /uninstall exited with $ccmExitCode. Uninstall should be continuing in the background."
	} catch {
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ERROR: ccmsetup.exe failed to start: $_"
		$ccmSetupError = $true
	}
} else {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] WARNING: [$CCMExePath] not found."
	$ccmSetupError = $true
	$ccmSetupMissing = $true
}

# If we can't start ccmsetup.exe, skip the log parsing.
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Waiting for MCM client uninstall to complete..."

# First, check if we couldn't start ccmsetup.
if ($ccmSetupMissing) {
	if ($svcCcmExec -And -Not [string]::IsNullOrEmpty($mcmVersion)) {
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] ccmsetup.exe does not exist, but ccmexec service and MCM client version still exist. This may impact reinstall. Removing Ccmexec service."
		try {
			# Stop service and wait for status to change.
			Stop-Service CcmExec -Force
			Start-Sleep -Seconds 5
			$svcCcmExec = Get-Service CcmExec
			$TimeoutDate = (Get-Date).AddMinutes(3)
			while ((Get-Date) -lt $TimeoutDate -And $svcCcmExec.Status -eq 'Running') {
				$svcCcmExec = Get-Service CcmExec
				Start-Sleep -Seconds 5
			}
			if ((Get-Date) -ge $TimeoutDate) {
				Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] WARNING: Timeout reached waiting for ccmexec service status change (waited 3 minutes). Current status: $($svcCcmExec.Status)"
			}
			sc.exe delete CcmExec
			Start-Sleep -Seconds 10
			$TimeoutDate = (Get-Date).AddMinutes(3)
			while ((Get-Date) -lt $TimeoutDate -And $svcCcmExec) {
				$svcCcmExec = Get-Service CcmExec -ErrorAction SilentlyContinue
				Start-Sleep -Seconds 5
			}
			if ((Get-Date) -ge $TimeoutDate) {
				Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] WARNING: Timeout reached waiting for ccmexec service to be removed (waited 3 minutes). Current status: $($svcCcmExec.Status)"
			}
		} catch {
			Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] FATAL ERROR: Failed to remove ccmexec service: $_"
			Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Exiting with return code -1"
			Stop-Transcript | Out-Null
			exit -1
		}
	}
} else {
	if (-Not $ccmSetupError) {
		# ccmsetup started successfully.
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
						Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] FATAL ERROR: MCM client uninstall unsuccessful. Ccmsetup exited with return code $returnCode (ccmsetup already running). Check ccmsetup.log for more info."
						Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Exiting with return code $returnCode."
						Stop-Transcript | Out-Null
						exit $returnCode
					} else {
						Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] FATAL ERROR: MCM client uninstall unsuccessful. Ccmsetup exited with return code $returnCode. Check ccmsetup.log for more info. Exiting with return code $returnCode."
						Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Exiting with return code $returnCode."
						Stop-Transcript | Out-Null
						exit $returnCode
					}
				}
			}

			Start-Sleep -Seconds 15
		}

		if ((Get-Date) -ge $TimeoutDate) {
			Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] WARNING: MCM client uninstall wait reached $SetupTimeoutMin minute timeout."
		}
	}

	Start-Sleep -Seconds 15

	# Additional validation after uninstall attempt
	# Wait up to $WmiTimeoutMin (5 minutes) for the SMS_Client WMI class to be removed.
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
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] WARNING: SMS_Client WMI class still exists with version [$($oSmsClient.ClientVersion)] after MCM client uninstall attempt. Waited $WmiTimeoutMin minutes."
	}

	$svcCcmExec = Get-Service CcmExec -ErrorAction SilentlyContinue
	if (-not $svcCcmExec) {
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Confirmed ccmexec service was successfully removed by ccmsetup."
	} else {
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] WARNING: CcmExec service still exists after MCM client uninstall attempt (current status: $($svcCcmExec.Status). This may prevent the reinstall script from triggering."
	}

	$mcmVersion = Get-MCMVersion
	if ([string]::IsNullOrEmpty($mcmVersion)) {
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] MCM Client Version registry entry was successfully wiped by ccmsetup."
	} else {
		Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] WARNING: MCM client version is still set as [$mcmVersion] in registry after uninstall attempt. This may prevent the reinstall script from triggering."
	}
}

if ($svcCcmExec -And -Not [string]::IsNullOrEmpty($mcmVersion)) {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] FATAL ERROR: ccmexec service and MCM client version both still exist after uninstall attempt. This will prevent the MCM client script from running on startup. Try running this script again."
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Exiting with return code -2"
	Stop-Transcript | Out-Null
	exit -2
}
	
if ($NoRestart) {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] -NoRestart given, skipping restart."
} else {
	Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Restarting in $($RestartSec / 60) minutes."
	shutdown.exe /r /t $RestartSec /c "SCCM client removal completed. Please restart the system. If the system is not restarted sooner, it will automatically restart in $($RestartSec / 60) minutes."
}
				
Write-Host "[$(Get-Date -f 'MM-dd-yyyy HH:mm:ss')] Exiting with return code 0"
Stop-Transcript | Out-Null
exit 0