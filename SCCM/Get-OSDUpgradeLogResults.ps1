# Exports Windows upgrade results parsed from the osdlogs folder.
param(
	[Parameter(Mandatory=$false)]
    [string] $LogsDir = '\\win.ad.jhu.edu\Data\osdlogs$\USS-ECI Upgrade Win11 25H2 x64 2026.02-V1.1',
	
	[Parameter(Mandatory=$true)]
	[string] $CSV
)

if (!(Test-Path $LogsDir)) {
	Write-Error "[$LogsDir] does not exist or is not accessible."
	exit 1
}

Write-Host "Searching [$LogsDir] for result directories..."
$results = Get-ChildItem -Path (Join-Path $LogsDir '\*') -Directory |
    ForEach-Object {
		

		try {
			$foldername = $_.FullName -split '\\' | Select -Last 1
			if (-Not $foldername) {
				Write-Warning "Problem parsing [$($_.FullName)]"
				continue
			}
		} catch {
			Write-Warning "Error parsing [$($_.FullName)]: $($_.Exception.Message)"
		}
		
		Write-Host "Found [$($_.FullName)] with folder name [$foldername]"
		
		if ($foldername -match "([^\.]+)\.([\d\.]+)@([\d\s\.]+)\.(.+)`$") {
			$systemName = $Matches[1]
			$logDate = $Matches[2]
			$logTime = $Matches[3]
			$logResult = $Matches[4]
		
			$diagfp = Join-Path $_.FullName 'setupdiagresults.log'
			if (-Not (Test-Path $diagfp )) {
				Write-Warning "[$diagfp] does not exist, not parsing"
			} else {
				$content = Get-Content $diagfp -Raw
				if ($content -match "Last Operation = (.+)") {
					$lastop = $Matches[1]
				} else {
					$lastop = $null
				}
				if ($content -match "Error = (.+)") {
					$lastop_error = $Matches[1]
				} else {
					$lastop_error = $null
				}
				if ($content -match "UpdateStartTime = (.+)") {
					$upgradeStart = $Matches[1]
				}
				if ($content -match "UpdateEndTime = (.+)") {
					$upgradeEnd = $Matches[1]
				}
				if ($content -match "RollbackStartTime = (.+)") {
					$rollbackStart = $Matches[1]
				}
				if ($content -match "RollbackEndTime = (.+)") {
					$rollbackEnd = $Matches[1]
				}
			}
			
			[PSCustomObject] @{
				'System' = $systemName
				'Result' = $logResult
				'Date' = $logDate
				'Time' = $logTime
				'Error' = $lastop_error
				'Last Op before Error' = $lastop
				'Upgrade Start Time' = $upgradeStart
				'Upgrade End Time' = $upgradeEnd
				'Rollback Start Time' = $rollbackStart
				'Rollback End Time' = $rollbackEnd
				'Log Path' = $_.FullName
			}
		} else {
			Write-Warning "Problem parsing folder name [$foldername] for [$($_.FullName)]"
		}
    }

if (-Not $results -Or $results.Count -eq 0) {
	Write-Warning "No results to export."
} else {
	Write-Host "* Exporting $($results.Count) results to [$CSV]..."
	# Add whether any items later upgraded successfully and attempt #.
	$sysResults = @{}
	foreach ($system in $results.System) {
		if (-Not $sysResults.ContainsKey($system)) {
			$sysResults[$system] = $results | where {$_.System -eq $system} | Sort Date,Time
		}
	}
	$results | Select 'System','Result','Date','Time',
		@{N="Eventual Success"; E={ 
			if ($_.Result -ne 'upgrade_success') {
				($sysResults[$_.System] | where {$_.Result -eq 'upgrade_success'} | Measure).Count -gt 0
			} else {
				''
			}
		}},
		@{N="Attempt"; E={ 
			if ($sysResults[$_.System] -is [array]) {
				[Array]::IndexOf($sysResults[$_.System], $_) + 1
			} else {
				1
			}
		}},
		'Error','Last Op before Error','Upgrade Start Time','Upgrade End Time','Rollback Start Time','Rollback End Time','Log Path' |
		Sort Date,Time |
		Export-CSV -NoTypeInformation $CSV
}