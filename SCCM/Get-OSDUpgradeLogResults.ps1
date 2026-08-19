# Exports Windows upgrade results parsed from the osdlogs folder.
param(
	[Parameter(Mandatory=$false)]
    [string] $LogsDir = '\\win.ad.jhu.edu\Data\osdlogs$\USS-ECI Upgrade Win11 25H2 x64 2026.02-V1.1',
	
	[Parameter(Mandatory=$true)]
	[string] $CSV
)

if (-Not (Test-Path $LogsDir) ) {
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
		
		Write-Host "Found [$($_.FullName)] with folder name [$foldername]. Parsing folder name and setup files..."
		
		if ($foldername -match "([^\.]+)\.([\d\.]+)@([\d\s\.]+)\.(.+)`$") {
			$systemName = $Matches[1]
			$logDate = $Matches[2]
			$logTime = $Matches[3]
			$logResult = $Matches[4]
		
			$diagfp = Join-Path $_.FullName 'setupdiagresults.log'
			if (-Not (Test-Path $diagfp )) {
				Write-Warning "Cannot find file: [$diagfp], skipping setup file parsing"
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
	Write-Host "** No results to export."
} else {
	Write-Host "** Exporting $($results.Count) results to [$CSV]..."
	# First sort by date & time.
	$sorted = @($results | Sort-Object Date, Time)
	
	# Add whether any items later upgraded successfully and attempt #.
	# Iterating $sorted preserves Date/Time order within each group automatically.
	$sysResults = @{}
	foreach ($row in $sorted) {
		if (-not $sysResults.ContainsKey($row.System)) {
			$sysResults[$row.System] = [System.Collections.ArrayList]::new()
		}
		[void]$sysResults[$row.System].Add($row)
	}

	# Exits early on first success hit — no redundant scanning per row later.
	$hasSuccess = @{}
	foreach ($sys in $sysResults.Keys) {
		$hasSuccess[$sys] = $false
		foreach ($row in $sysResults[$sys]) {
			if ($row.Result -eq 'upgrade_success') {
				$hasSuccess[$sys] = $true
				break
			}
		}
	}

	# Uses object reference as hashtable key.
	$attemptLookup = [System.Collections.Hashtable]::new($sorted.Count)
	foreach ($sys in $sysResults.Keys) {
		$group = $sysResults[$sys]
		for ($i = 0; $i -lt $group.Count; $i++) {
			$attemptLookup[$group[$i]] = $i + 1
		}
	}

	$sorted |
		Select-Object 'System', 'Result', 'Date', 'Time',
			@{N = 'Eventual Success'; E = {
				if ($_.Result -ne 'upgrade_success') { $hasSuccess[$_.System] } else { '' }
			}},
			@{N = 'Attempt'; E = { $attemptLookup[$_] }},
			'Error', 'Last Op before Error', 'Upgrade Start Time', 'Upgrade End Time',
			'Rollback Start Time', 'Rollback End Time', 'Log Path' |
		Export-CSV -NoTypeInformation $CSV
}