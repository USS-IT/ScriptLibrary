# MJC 3-10-22

# Searchbase for Get-ADComputer
$_SB = "OU=HSA-SCT,OU=Classrooms,OU=HOPKINS,DC=win,DC=ad,DC=jhu,DC=edu"
# $_SB = "OU=HSA-HSAT,OU=Depts,OU=HSA-SCT,OU=Classrooms,OU=HOPKINS,DC=win,DC=ad,DC=jhu,DC=edu"

# Filename for results (always saved with ".csv" file extension).
$_DESTFN = "localadmins"
# Path to save results.
$_SAVE_PATH = "."
# Path to archived results.
$_ARCHIVE_PATH = ".\Archive"
# Number of days of historical records to save.
$_ARCHIVE_DAYS=30
# Path to log files.
$_LOG_PATH=".\Logs"

# Prune old log files and start logging
if ($_LOG_PATH) {
	if (-Not (Test-Path $_LOG_PATH)) {
		New-Item -ItemType Directory -Force -Path $_LOG_PATH
	}
	Get-ChildItem "${_LOGPATH}\Export-LocalAdmins_*.log" -ErrorAction SilentlyContinue | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-30) } | Remove-Item -Force
	Start-Transcript -Path "${_LOG_PATH}\Export-LocalAdmins_$(get-date -f yyyy-MM-dd).log"
}

function WriteLog($msg) {
	$stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
	Write-Host "[${stamp}] ${msg}"
}

# Load file which contains the Get-LocalAdmins function.
. .\Get-LocalAdmins.ps1

# Import previous results.
$prev_results = $null
$dest_fp = "${_SAVE_PATH}\${_DESTFN}.csv"
if (Test-Path -Path $dest_fp -PathType Leaf) {
	$prev_results = Import-CSV $dest_fp
}

# For filtering out default local admin groups and the built-in administrator
$defaultAdmins = Get-LocalGroupMember Administrators | where {$_.ObjectClass -eq "Group"} | Select -ExpandProperty Name
$localAdminName = (Get-LocalUser | where {$_.SID -like "S-1-5-*-500" }).Name | Select -First 1

# Comment out the line below to disable filtering the default admin users & groups, or to make your own list
# $defaultAdmins = @()
WriteLog("Local Admin Name: $localAdminName")
WriteLog("Default Local Admin Groups: {0}" -f ($defaultAdmins -join ', '))

# Get all enabled computers starting at searchbase, return only those with different local admins than current machine
WriteLog("Querying all enabled computers at [$_SB]...")
$comps = Get-ADComputer -searchbase $_SB -Filter {(Enabled -eq $True)} | Select Name, DistinguishedName | Foreach-Object {
	$name = $_.Name
	$online = $False
	$errorcaught = $False
	$allAdmins = ""
	$uniqueAdmins = ""
	WriteLog("Checking local admins for [$name]...")
	If ((Test-Connection -ComputerName $name -Count 1 -Quiet) -OR (Test-Connection -ComputerName $name -Count 3 -Quiet)) {
		$online = $True
		try {
			$allAdmins = Get-LocalAdmins($name)
			$uniqueAdmins = ($allAdmins | where {$_.SAMAccountName -notin $defaultAdmins -And $_.Name -ne $localAdminName} | Select -ExpandProperty SAMAccountName) -join "; "
			$allAdmins = ($allAdmins | Select -ExpandProperty SAMAccountName) -join "; "
		} catch {
			WriteLog("ERROR retreiving local admins for [$name]")
			$allAdmins = "<ERROR RETRIEVING LOCAL ADMINS>"
			$uniqueAdmins = $allAdmins
		}
	}
	if ((-Not ($online -OR $errorcaught)) -And $prev_results) {
		# Select from previous results
		$comp = $prev_results | where { $_.Name -eq $name } | Select -First 1
		if ($comp -And $comp.AllLocalAdmins) {
			WriteLog("Using previous result for [$name]")
			$allAdmins = $comp.AllLocalAdmins
			$uniqueAdmins = $comp.UniqueLocalAdmins
		}
	}
	[PSCustomObject][ordered]@{ 
		Name = $name
		Online = $online
		UniqueLocalAdmins = $uniqueAdmins
		AllLocalAdmins = $allAdmins
		DistinguishedName = $_.DistinguishedName
	}
}
if (-Not $comps) {
    WriteLog ("No results, skipping saving to [$dest_fp]")
} else {
    if (Test-Path -Path $dest_fp -PathType Leaf) {
        # Archive previous results.
	    if ($_ARCHIVE_DAYS -And $_ARCHIVE_DAYS -gt 0 -And $_ARCHIVE_PATH) {
            $fp = "{_ARCHIVE_PATH}\${_DESTFN}_$(get-date -f yyyy-MM-dd).csv"
            WriteLog("Archiving previous results to [$fp] and pruning older than $_ARCHIVE_DAYS days")
		    if (-Not (Test-Path $_ARCHIVE_PATH)) {
			    New-Item -ItemType Directory -Force -Path $_ARCHIVE_PATH
		    }
		    Move-Item -Path $dest_fp -Destination $fp -Force
		    Get-ChildItem "${_ARCHIVE_PATH}\${_DESTFN}_*.csv" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$_ARCHIVE_DAYS) } | Remove-Item -Force
	    }
    }
    WriteLog ("Saving $($comps.Count) results to [$dest_fp]")
    $comps | Export-CSV -NoTypeInformation $dest_fp
}