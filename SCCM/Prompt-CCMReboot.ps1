<#
    .SYNOPSIS
    Tells the CCM client to prompt for a restart.
    
	.DESCRIPTION
    Tells the CCM client to prompt for a restart.
	
	.PARAMETER RestartService
	Restarts the CCM service. This is needed if called outside CCM. Takes about 30-60 seconds after restart for prompt.
	
	.PARAMETER NonMandatory
    Sets reboot prompt to Non-Mandatory (no deadline).
	
    .NOTES
	mcarras8 8-25-25
#>
param (
	[switch] $RestartService,
	[switch] $NonMandatory
)
# Get the current reboot status.
$ocim = Invoke-CimMethod -Namespace root/ccm/ClientSDK -ClassName CCM_ClientUtilities -MethodName DetermineIfRebootPending

# Force a reboot prompt from MCM Client, but only if it hasn't already been prompted. 
if ($ocim.RebootPending) {
	Write-Verbose "[Prompt-CCMReboot.ps1] Already pending reboot"
} else {
	if ($NonMandatory) {
		$rebootTime = 0
	} else {
		$rebootTime = [DateTimeOffset]::Now.ToUnixTimeSeconds()
	}
	Write-Verbose "[Prompt-CCMReboot.ps1] Setting registry for CCM pending reboot"
	New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'RebootBy' -Value $rebootTime -PropertyType QWord -Force -ea SilentlyContinue;
	New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'RebootValueInUTC' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
	New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'NotifyUI' -Value 1 -PropertyType DWord -Force -ea SilentlyContinue;
	New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'HardReboot' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
	New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'OverrideRebootWindowTime' -Value 0 -PropertyType QWord -Force -ea SilentlyContinue;
	New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'OverrideRebootWindow' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
	New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'PreferredRebootWindowTypes' -Value @("4") -PropertyType MultiString -Force -ea SilentlyContinue;
	New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData' -Name 'GraceSeconds' -Value 0 -PropertyType DWord -Force -ea SilentlyContinue;
	
	if ($RestartService) {
		Write-Verbose "[Prompt-CCMReboot.ps1] Restarting ccmexec"
		Restart-Service ccmexec -force
	}
}
