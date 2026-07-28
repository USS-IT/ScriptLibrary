<#
.SYNOPSIS
    Adds temporary DNS server(s) to primary network adapter or resets it
.DESCRIPTION
    This script can:
	1) Add temporary DNS servers to the primary network adapter
	2) Remove temporary DNS servers from the primary network adapter / reset back to DHCP DNS
.PARAMETER DNSServers
	DNS Server(s) to add. These must all be the same type (IPv4 or IPv6). Default: 8.8.8.8
.PARAMETER DNSOverride
	Override the DNS instead of adding to it.
.PARAMETER StaticDNSJson
	Path to store any previously set static DNS settings. Default: C:\TEMP\SwitchDNS.json
.PARAMETER Mode
	Sets the change mode: Toggle, Add, or Reset. Default: Toggle
.PARAMETER ResetOnStartup
	When changing DNS, add a scheduled task to reset it once on startup.
.PARAMETER Notify
	Invokes an event-based task to notify the current user.
.PARAMETER LogDir
	Directory for logging.
.NOTES
	Author: Matt Carras (mcarras8)
	Created: 07-07-2026
#>
#Requires -RunAsAdministrator
param(
	[IPAddress[]] $DNSServers = [IPAddress]'8.8.8.8',
		
	[switch] $DNSOverride,
	
	[string] $StaticDNSJson = "C:\TEMP\SwitchDNS.json",
	
	[ValidateSet('Toggle','Add','Reset')]
	[string] $Mode='Toggle',
	
	[switch] $ResetOnStartup,
		
	[switch] $Notify,
	
	[string] $LogDir = "C:\USS\Logs\Packages\SwitchDNS"
)

# Task name for the on startup reset task.
$ResetTaskName = "USS-SwitchDNS-Reset"

try {
	$_scriptName = Split-Path -Leaf $PSCommandPath
} catch {
	$_scriptName = "Switch-DNS.ps1"
}
$LogPath = "$LogDir\$($_scriptName).log"

# Create the log path if it doesn't already exist
if (-not (Test-Path $LogDir -PathType Container)) {
	$null = New-Item -Path $LogDir -ItemType Directory
}
Start-Transcript $LogPath -Force

$AddressFamily = $DNSServers.AddressFamily | Select -Unique
if (($AddressFamily | Measure).Count -gt 1) {
	Write-Error "All IP Addresses must be the same type (IPv4 or IPv6)."
	Stop-Transcript | Out-Null
	[System.Environment]::Exit(-1)
}
if (($AddressFamily | Select -First 1) -eq 'InterNetworkV6') {
	$AddressFamily = 'IPv6'
} else {
	$AddressFamily = 'IPv4'
}

$exitCode = 0
try {
	# Get all physical network adapters
	$adapters = Get-NetAdapter -Physical

	foreach ($adapter in $adapters) {
		Write-Host "Checking [$($adapter.InterfaceDescription)] with status [$($adapter.Status)]"
		$ifIndex = $adapter.ifIndex
		$currentDns = (Get-DnsClientServerAddress -InterfaceIndex $ifIndex -AddressFamily $AddressFamily).ServerAddresses
		Write-Host "Current $AddressFamily DNS: $($currentDns -join ',')"
		
		if ($Mode -ne 'Add' -And $currentDns -contains $DNSServers.IPAddressToString) {
			$prevDNS = $null
			Write-Host "DNS currently includes $($DNSServers.IPAddressToString -join ','). Checking if we have any previously saved static DNS settings..."
			if ((Test-Path $StaticDNSJson -PathType Leaf)) {
				$prevDNS = Get-Content -Path $StaticDNSJson
				if (-Not [string]::IsNullOrEmpty($prevDNS)) {
					$prevDNS = $prevDNS | ConvertFrom-Json | where {$DNSServers.IPAddressToString -notcontains $_}
				}
			}
			
			if ($prevDNS -ne $null -And -Not [string]::IsNullOrWhitespace(($prevDNS | Select -First 1))) {
				Write-Host "Reverting to previous DNS [$($prevDNS -join ',')]"
				Set-DnsClientServerAddress -InterfaceIndex $ifIndex -ServerAddresses $prevDNS -ErrorAction Stop
				Clear-Content $StaticDNSJson
			} else {
				Write-Host "Reverting to DHCP DNS"
				Set-DnsClientServerAddress -InterfaceIndex $ifIndex -ResetServerAddresses -ErrorAction Stop
			}
			
			# Unregister reset task if it exists.
			if ((Get-ScheduledTask -TaskName $ResetTaskName -ErrorAction SilentlyContinue)) {
				Unregister-ScheduledTask -TaskName $ResetTaskName -Confirm:$false
			}
			
			if ($Notify) {
				Write-Host "Sending notification"
				Write-EventLog -LogName 'USS-EventLog' -Source 'Notify-SwitchDNS-Reverted' -EventID 1000 -EntryType Information -Message 'Notify of successful DNS change'
			}
		} elseif ($Mode -ne 'Reset' -And $adapter.Status -eq 'Up') {
			Write-Host "Adapter is currently up, and DNS does not include $($DNSServers.IPAddressToString -join ',')."
		
			# Check if we have any statically added DNS already.
			$key = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\$($adapter.InterfaceGuid)"
			$staticDNS = Get-ItemProperty $key -ErrorAction Stop | Select -ExpandProperty NameServer
			if (-Not [string]::IsNullOrEmpty($staticDNS)) {
				Write-Host "Current static DNS: $staticDNS"
				$dir = Split-Path -Parent $StaticDNSJson -ErrorAction Stop
				if (-not (Test-Path $dir -PathType Container)) {
					$null = New-Item -Path $dir -ItemType Directory -ErrorAction Stop
				}
				$staticDNS = $staticDNS -split ',' | where {$DNSServers.IPAddressToString -notcontains $_}
				if (($staticDNS | Measure).Count -gt 0 -And -Not [string]::IsNullOrWhitespace(($staticDNS | Select -First 1))) {
					Write-Host "Saving static DNS to [$StaticDNSJson]: $($staticDNS -join ',')"
					Set-Content -Path $StaticDNSJson -Value ($currentDns | ConvertTo-Json) -ErrorAction Stop
				}
			}

			if ($DNSOverride) {
				$newDNS = $DNSServers.IPAddressToString
			} else {
				$newDNS = @($currentDns) + $DNSServers.IPAddressToString
			}
			Write-Host "Setting DNS servers to: $($newDNS -join ',')"
			Set-DnsClientServerAddress -InterfaceIndex $ifIndex -ServerAddresses $newDNS -ErrorAction Stop
			
			if ($ResetOnStartup) {				
				if ((Get-ScheduledTask -TaskName $ResetTaskName -ErrorAction SilentlyContinue)) {
					Write-Host "Scheduled task [$ResetTaskName] already exists. Not creating."
				} else {
					Write-Host "Creating AtStartup scheduled task to reset DNS..."
					$scriptPath = $PSCommandPath
					If ([string]::IsNullOrEmpty($scriptPath)) {
						$scriptPath = "C:\USS\Scripts\Admin\SwitchDNS\Switch-DNS.ps1"
					}
				
					$action = New-ScheduledTaskAction `
						-Execute "powershell.exe" `
						-Argument "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Mode `"Reset`""
					
					$trigger = New-ScheduledTaskTrigger -AtStartup
					
					$principal = New-ScheduledTaskPrincipal `
						-UserId "SYSTEM" `
						-LogonType ServiceAccount `
						-RunLevel Highest
						
					$settings = New-ScheduledTaskSettingsSet `
						-AllowStartIfOnBatteries `
						-DontStopIfGoingOnBatteries `
						-StartWhenAvailable `
						-ExecutionTimeLimit (New-TimeSpan -Hours 2) `
						-Hidden
			
					$task = Register-ScheduledTask `
						-TaskName $ResetTaskName `
						-Action $action `
						-Trigger $trigger `
						-Principal $principal `
						-Settings $settings
						-Description "Reverts DNS changes from Switch-DNS.ps1."
						-Force

					Write-Host "Scheduled task [$($task.TaskName)] created successfully."
				}
			}
			
			if ($Notify) {
				Write-Host "Sending notification"
				Write-EventLog -LogName 'USS-EventLog' -Source 'Notify-SwitchDNS-Changed' -EventID 1000 -EntryType Information -Message 'Notify of successful DNS change'
			}
		}
	}
} catch {
	Write-Error $_
	$exitCode = 1
	if ($Notify) {
		Write-Host "Sending notification"
		Write-EventLog -LogName 'USS-EventLog' -Source 'Notify-SwitchDNS-Failure' -EventID 1000 -EntryType Information -Message 'Notify of failed DNS change'
	}
}

Write-Host "Flushing DNS cache..."
Clear-DnsClientCache

Write-Host "Done."

Stop-Transcript | Out-Null
[System.Environment]::Exit($exitCode)
