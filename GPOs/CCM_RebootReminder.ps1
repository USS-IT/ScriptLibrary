<#
	.SYNOPSIS
	Uses the Configuration Manager Client to either initiate a snoozable automatic restart (mandatory) or prompt for one (non-mandatory).
	
	.DESCRIPTION
	Uses the Configuration Manager Client to either initiate a snoozable automatic restart (mandatory) or prompt for one (non-mandatory). This should be run from a non-writable directory as an Admin.

	.PARAMETER RestartAlertDays
	Number of days before an alert is displayed. Default is 7.
	
	.PARAMETER LastRunFile
	A file which contains the last date the reminder was shown. Give an empty string to disable. Default: C:\TEMP\RestartReminderLastRun
	
	.PARAMETER LastRunWaitHours
	The amount of time the script will wait after the date read from the LastRunFile, if it exists. Default: 24 hours
	
	.PARAMETER RestartService
	Restarts the CCM service. This is needed if called outside CCM. Takes about 30-60 seconds after restart for prompt.
	
	.PARAMETER NonMandatory
    Sets reboot prompt to Non-Mandatory (no deadline).
	
	.NOTES
	Must be run under the user's context.
	
	Created: 3-23-23
	Author: mcarras8
#>
[cmdletbinding(DefaultParametersetName='None')]
param(
	[Parameter(Mandatory=$false)]
	[int]$RestartAlertDays=7,
	
	[Parameter(Mandatory=$false)]
	[AllowEmptyString()]
	[string]$LastRunFile="C:\TEMP\RestartReminderLastRun",
	
	[Parameter(Mandatory=$false)]
	[int]$LastRunWaitHours=24,
	
	[Parameter(Mandatory=$false)]
	[switch]$RestartService,
	
	[Parameter(Mandatory=$false)]
	[switch]$NonMandatory
)

try {
	$doProcess = $true
	# Check the LastRun file if it's available.
	if (-Not [string]::IsNullOrEmpty($LastRunFile) -And (Test-Path $LastRunFile -PathType Leaf)) {
		# Only show if the script was last run over $LastRunWaitHours hours ago (default: every 24 hours).
		$lastRunDate = (Get-Content $LastRunFile) -as [datetime]
		Write-Verbose ("[CCM_RebootReminder.ps1] lastRunDate=$lastRunDate, Hours Diff={0}" -f ((Get-Date) - $lastRunDate).TotalHours)		
		if ($lastRunDate -is [datetime] -And ((Get-Date) - $lastRunDate).TotalHours -le $LastRunWaitHours) {
			Write-Verbose "[CCM_RebootReminder.ps1] Skipping due to lastRunDate diff less than LastRunWaitHours"
			$doProcess = $false
		}
	}
	if ($doProcess) {
		# Check if we already have a pending reboot.
		$ocim = Invoke-CimMethod -Namespace root/ccm/ClientSDK -ClassName CCM_ClientUtilities -MethodName DetermineIfRebootPending
		if ($ocim.RebootPending) {
			Write-Verbose "[CCM_RebootReminder.ps1] CCM already has pending reboot, aborting"
		} else {
			# Grab uptime from WMI.
			$uptimeDays = ((Get-Date) - [Management.ManagementDateTimeConverter]::ToDateTime((Get-WmiObject Win32_OperatingSystem | Select -ExpandProperty LastBootUpTime))) | Select -ExpandProperty Days
			if ($uptimeDays -ge $RestartAlertDays) {
				if ($NonMandatory) {
					$rebootTime = 0
				} else {
					$rebootTime = [DateTimeOffset]::Now.ToUnixTimeSeconds()
				}
				Write-Verbose "[CCM_RebootReminder.ps1] Setting registry for CCM pending reboot"
				New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'RebootBy' -Value $rebootTime -PropertyType QWord -Force -ea SilentlyContinue;
				New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'RebootValueInUTC' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
				New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'NotifyUI' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
				New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'HardReboot' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
				New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'OverrideRebootWindowTime' -Value 0 -PropertyType QWord -Force -ea SilentlyContinue;
				New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'OverrideRebootWindow' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
				New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'PreferredRebootWindowTypes' -Value @("4") -PropertyType MultiString -Force -ea SilentlyContinue;
				New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'GraceSeconds' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
				
				if ($RestartService) {
					Write-Verbose "[CCM_RebootReminder.ps1] Restarting ccmexec service"
					Restart-Service ccmexec -force
				}
				
				# Add to the current date & time to the lastrun file (if it's set).
				if (-Not [string]::IsNullOrEmpty($LastRunFile)) {
					Write-Verbose "[CCM_RebootReminder.ps1] Updating LastRunFile [$LastRunFile]"
					Set-Content -Path $LastRunFile -Value (Get-Date)
				}
			} else {
				Write-Verbose "[CCM_RebootReminder.ps1] No reminders due to uptime [$uptimeDays] < [$RestartAlertDays] alertdays"
			}
		}
	}
} catch {
	Write-Error $_
}