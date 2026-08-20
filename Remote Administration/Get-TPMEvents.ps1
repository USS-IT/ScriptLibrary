# Get TPM Secure Boot Certificate related events from the Windows Event Log for remote systems.
# 1034 - Secure Boot dbx updated successfully
# 1801, 1808 - Updated Secure Boot CA available

[CmdletBinding(DefaultParameterSetName='ComputerName')]
param(
	[Parameter(Mandatory=$true, ParameterSetName = 'ComputerName')]
    [string] $ComputerName,
	
	[Parameter(Mandatory=$true, ParameterSetName = 'ComputerName')]
    [datetime] $Date,
	
	[Parameter(Mandatory=$true, ParameterSetName = 'SystemSets')]
    [PSObject[]] $SystemSets,
	
	[Parameter(Mandatory=$false, ParameterSetName = 'ComputerName')]
	[Parameter(Mandatory=$false, ParameterSetName = 'SystemSets')]
	[switch]   $NoPause
)

if (-Not $SystemSets) {
	$_SystemSets = [PSCustomObject]@{ 
		'System' = $ComputerName
		'Date' = $Date -as [datetime]
	}
} else {
	$_SystemSets = $SystemSets
}

$result1034Counts = @{
	'upgrade_success' = 0
	'upgrade_success_total' = 0
	'rollback' = 0
	'rollback_total' = 0
	'noUpg' = 0
	'noUpg_total' = 0
}

$queried = @{}
foreach ($system in $_SystemSets) {
	$systemName = $system.System
	
	if (-Not (Test-Connection -ComputerName $systemName -Count 1 -Quiet) -And 
		-Not (Test-Connection -ComputerName $systemName -Count 1 -Quiet) -And 
		-Not (Test-Connection -ComputerName $systemName -Count 1 -Quiet)) {
		
		Write-Warning "'$systemName' did not respond to any ping attempts. Skipping."
		continue
	}
	
	$skipCount = $false
	if ($queried.ContainsKey($systemName)) {
		Write-Host "'$systemName' queried previously"
		
		if (-Not [string]::IsNullOrEmpty($system.Result)) {
			if ($system.Result -in $queried[$systemName].Results) {
				$skipCount = $true
			} else {
				$queried[$systemName].Results.Add($system.Result)
			}
		}
	} else {
		try {
			if ($system.Result) {
				Write-Output "'$systemName' appears to be online. Logged result was [$($system.Result)]. Querying..."
			} else {
				Write-Output "'$systemName' appears to be online. Querying..."
			}
		
			$systemDate = $system.Date -as [datetime]
		
			# Build the 24-hour window for the target date
			$startTime = $systemDate.Date                             # 00:00:00
			$endTime   = $systemDate.Date.AddDays(1).AddTicks(-1)     # 23:59:59.9999999

			$filterHash = @{
				LogName      = "System"
				ProviderName = "Microsoft-Windows-TPM-WMI"
				#Id           = @(1808, 1801, 1034)
				StartTime    = $startTime
				EndTime      = $endTime
			}
	
			$events = Get-WinEvent -ComputerName $systemName `
								   -FilterHashtable $filterHash `
								   -ErrorAction Stop
			$queried[$systemName] = @{
				'Results' = New-Object -TypeName "System.Collections.ArrayList"
				'Events' = $events
			}
			$queried[$systemName].Results.Add($system.Result)
		
			if ($events.Count -eq 0) {
				Write-Output "No entries matching filter were found on '$systemName' for $($systemDate.ToString('yyyy-MM-dd'))."
			} else {
				Write-Output "Found $($events.Count) event(s) on '$systemName' for $($systemDate.ToString('yyyy-MM-dd')):"
				$events | Select-Object TimeCreated,
										Id,
										@{N="Source"; E={$_.ProviderName}},
										@{N="Log";    E={$_.LogName}},
										LevelDisplayName,
										Message |
						  Format-List
			}
		} catch [System.Exception] {
			if ($_.Exception.Message -like "*No events were found*") {
				Write-Output "ERROR: No entries matching filter were found on '$systemName' for $($systemDate.ToString('yyyy-MM-dd'))."
			} else {
				Write-Error "Failed to query '$systemName': $($_.Exception.Message)"
			}
		}
			
		if (-Not $skipCount) {
			if ($system.Result -eq 'upgrade_success') {
				$result1034Counts['upgrade_success_total']++
				if ($events -And 1034 -in $events.Id) {
					$result1034Counts['upgrade_success']++
				}
			} elseif ($system.Result -match 'rollback') {
				$result1034Counts['rollback_total']++
				if ($events -And 1034 -in $events.Id) {
					$result1034Counts['rollback']++
				}
			} elseif ($system.Result -match 'noUpg') {
				$result1034Counts['noUpg_total']++
				if ($events -And 1034 -in $events.Id) {
					$result1034Counts['noUpg']++
				}
			}
		}
	}
}

if ($SystemSets) {
	Write-Host "Event ID 1034 and Overall Totals"
	$result1034Counts
}

if (-Not $NoPause) {
	$_ = Read-Host ":Press any key to exit"
}