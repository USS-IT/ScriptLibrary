<#
	.SYNOPSIS
	Move objects matching name filters to given OUs.
	
	.DESCRIPTION
	Move objects matching name filters to given OUs. Requires RSAT tools.
	
	.NOTES
	Looks in $OU_MapSearchbase for current AD structure to use for name filtering.
	Any OUs starting in $OU_MapSearchbase will be added to name filtering in the form of "OUName-*".
	IE, if it finds an OU named "USS-IT" then it will search for computers starting with "USS-IT-".
	You may override this behavior using the $destOU_map hashtable.
	$destOU_map hash table:
	{
		key = destination OU distinguishedname.
		"inc" = array of partial computer names to include.
		"ex" = optional array of partial computer names to exclude.
	}
	
	Created: 10-18-22
	Author: mcarras8
	
	Changelog
	03-14-25 - MJC - Standardized documentation and uploaded to github
	04-16-24 - MJC - Added emailed error reports and additional error handling
	01-03-23 - MJC - Added CDI mapping for USS-OMA OU
#>
<# --EXAMPLES--
	$DC = "DC=win,DC=ad,DC=jhu,DC=edu"
	$OU_MoveDestRoot = "OU=Computers,OU=USS,$DC"
	$destOU_map = @{
		"OU=USS-SC,$OU_MoveDestRoot"	= @{ "inc"=@("USS-CSC-","HW-SC-","HW-CSC-") }
		"OU=USS-SEAM,$OU_MoveDestRoot"	= @{ "inc"=@("USS-SEAM-","HW-SEAM-","SEAM-")
											 "ex"=@("SEAM-DC-") }
#>

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs"
$LOGFILE_PREFIX = "Move-ADComputers"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 365

# Email configuration for reports
$EMAIL_SMTP = 'smtp.johnshopkins.edu'
# If filled out, send error reports
$EMAIL_ERROR_REPORT_FROM = 'USS IT Services <ussitservices@jhu.edu>'
# Can be string or array of strings.
$EMAIL_ERROR_REPORT_TO = @('mcarras8@jhu.edu','ussitservices@jhu.edu')

# Domain controller CN.
$DC = "DC=win,DC=ad,DC=jhu,DC=edu"
# Search the given OU for results and add them to the OU map if they don't already exist.
$OU_MapSearchbase = "OU=Computers,OU=USS,$DC"
# Name Filter used for searching OUs (wildcard accepted)
$OU_NameFilter = "USS-*"
# Default destination root OU
$OU_MoveDestRoot = "OU=Computers,OU=USS,$DC"
# Array of OUs to search for computer objects that need to be moved
$OU_MoveSearchbases = @("OU=USS-XX,OU=Computers,OU=USS,$DC","OU=USS-Retired,OU=Computers,OU=USS,$DC","OU=USS-LNR,OU=Computers,OU=USS,$DC")

# Entries here will override the entries added when searching $OU_Searchbase.
$destOU_map = @{
	"OU=USS-DMG,OU=USS-DMC,$OU_MoveDestRoot" = @{ "inc"=@("USS-DGL-") }
	"OU=USS-SHWB,$OU_MoveDestRoot"		= @{ "inc"=@("USS-HW-","HW-HW-") }
	"OU=USS-HD,$OU_MoveDestRoot" 		= @{ "inc"=@("USS-DO-","USS-RDO") }
	"OU=USS-LDL,$OU_MoveDestRoot" 		= @{ "inc"=@("USS-LDL-","HW-LDL-","HW-CP-") }
	"OU=USS-OMA,$OU_MoveDestRoot" 		= @{ "inc"=@("USS-OMA-","HW-OMA-","USS-CDI-") }
    "OU=USS-PRS,$OU_MoveDestRoot" 		= @{ "inc"=@("USS-PS-","USS-PRS-") }
	"OU=USS-ROC,$OU_MoveDestRoot" 		= @{ "inc"=@("USS-ROC-","HW-ROTC-","ROTC-") }
	"OU=USS-SEAM,$OU_MoveDestRoot" 		= @{ "inc"=@("USS-SEAM-","HW-SEAM-","SEAM-") }
	"OU=USS-ST,$OU_MoveDestRoot" 		= @{ "inc"=@("USS-ST-","USS-SSA","HW-ST-") }
	"OU=USS-Testing,$OU_MoveDestRoot" 	= @{ "inc"=@("USS-ITT-","HW-HSAT-") }
}

# Entries here are aliases for OU names.
# E.x. "HW-RL" is an alias of "USS-RL".
<#
$OUAliases = @{
	"USS-SEAM" = @("HW-SEAM-","SEAM-"),
	"USS-RL" = @("HW-RL-")
}
#>
# -- END CONFIGURATION --

# -- START --

# Rotate log files
if ($LOGFILE_ROTATE_DAYS -is [int] -And $LOGFILE_ROTATE_DAYS -gt 0) {
	Get-ChildItem "${LOGFILE_PATH}\${LOGFILE_PREFIX}_*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$LOGFILE_ROTATE_DAYS) } | Remove-Item -Force
}

# Start logging
$_logfilepath = "${LOGFILE_PATH}\${LOGFILE_PREFIX}_$(get-date -f yyyy-MM-dd).log"
Start-Transcript -Path $_logfilepath -Append

$_scriptName = split-path $PSCommandPath -Leaf
$errorCount = 0

# Add additional results from current AD structure.
foreach ($ou in Get-ADOrganizationalUnit -Searchbase $OU_MapSearchbase -Filter "Name -LIKE '$OU_NameFilter'") {
	if (-Not [string]::IsNullOrWhitespace($ou.DistinguishedName) -And -Not $destOU_map.ContainsKey($ou.DistinguishedName)) {
		$destOU_map.Add($ou.DistinguishedName, @{ "inc"=@(($ou.Name + '-')) })
	}
}

# Get all computers in target searchbase.
try {
	$comps = $OU_MoveSearchbases | % { Get-ADComputer -SearchBase $_ -Properties distinguishedname -Filter {Enabled -eq $true} }
} catch {
	Write-Error $_
	$errorCount++
}

# Loop over each filter to see which apply, adding them to a move queue.
# This makes sure we're not moving the same computer twice.
$moveQueue = @{}
foreach ($ou in $destOU_map.Keys) {
	# Combines multiple partial names into text for -Filter.
	$filter = ($destOU_map[$ou]["ex"] | foreach {if (-Not [string]::IsNullOrWhitespace($_)){"`$_.Name -NOTLIKE `"$_*`""}}) -join " -AND "
	$filter2 = ($destOU_map[$ou]["inc"] | foreach {if (-Not [string]::IsNullOrWhitespace($_)){"`$_.Name -LIKE `"$_*`""}}) -join " -OR "
	if (-Not $filter) { $filter = $filter2 } 
	elseif ($filter2) { $filter += " -AND ($filter2)" }
	Write-Verbose "OU: [$ou], Filter: [$filter]"
	if (-Not [string]::IsNullOrWhitespace($filter)) {
		# Loop over each computer using dynamic Where-Object filters.
		foreach ($comp in ($comps | Where-Object -FilterScript ([scriptblock]::create($filter)))) {
			# If computer is not already in target OU, add it to moveQueue.
			if ($ou -ne $comp.DistinguishedName.Substring($comp.DistinguishedName.IndexOf('OU='))) {
				try {
					$validatedOU = Get-ADOrganizationalUnit $ou
					$moveQueue[$comp.distinguishedname] = $validatedOU.DistinguishedName
				} catch {
					Write-Error $_
					Write-Error "** Invalid OU: $ou"
					$errorCount++
				}
			}
		}
	}
}

# Loop over the results in the moveQueue to move each computer.
# Debug: Use -WhatIf, don't actually move anything.
foreach ($dn in $moveQueue.Keys) {
	$targetPath = $moveQueue[$dn]
	if (-Not [string]::IsNullOrWhitespace($targetPath)) {
		try {
            Write-Host "[Move-ADComputers.ps1] Moving [$dn] to [$targetPath]" 
			Move-ADObject -Identity $dn -TargetPath $targetPath
		} catch {
			Write-Error $_
			$errorCount++
		}
	}
}

Write-Host("[0] Errors encountered: {1}" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $errorCount)

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

if (-Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP)) {  
    # Email out notifications of any errors.
    if ($errorCount -gt 0 -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1)))	{
		$emailParams = @{
			From = $EMAIL_ERROR_REPORT_FROM
			To =  $EMAIL_ERROR_REPORT_TO
			Subject = "Errors from $_scriptName"
			Body = "There were [$errorCount] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See attached logfile for more details."
			Priority = "High"
			DeliveryNotificationOption = @("OnSuccess", "OnFailure")
			SmtpServer = $EMAIL_SMTP
		}
        try {
			Send-MailMessage -Attachments $_logfilepath @emailParams
		} catch {
			Write-Error $_
			$mailParams.Body = "There were [$errorCount] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See [$_logfilepath] for more details."
			Send-MailMessage @emailParams
		}
        Write-Host("[{0}] Emailed error report to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($EMAIL_ERROR_REPORT_TO -join ", "))
    }
}
