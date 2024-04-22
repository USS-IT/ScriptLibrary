# Outputs information on EOL computers from AD.
# Uses attributes synced from Sync-ADWithSOR.ps1.
#
# Requires: RSAT AD Tools
#
# MJC 4-22-24

# Build number to check
$eolver=19044
# CSV Filename to output results
$csvFilename="USS EOL Computers $eolver or lower.csv"
# Fields to output to CSV file, in order
$csvfields = @("Name","OperatingSystemVersion","Contact User","UserType","extensionAttribute9","Link","extensionAttribute1","extensionAttribute2","extensionAttribute3","extensionAttribute10")

# The URL to lookup by the value in extensionAttribute1.
$sorurl = "https://jh-uss.snipe-it.io/hardware/bytag?assetTag="
# The searchbase to search for matching computers in AD.
$searchbase = "OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu"
# AD Computer Properties/Attributes used in logic. Assumes the following:
# extensionAttribute1 = asset tag
# extensionAttribute2 = assigned user by userprinciplname (used first for "Contact User")
# extensionAttribute5 = System Form Factor (Laptop, Tablet, etc.)
# extensionAttribute10 = LastLogonUser from query or SCCM export (used second for "Contact User" if no assigned user)
# extensionAttribute3 = Primary Users from SCCM export, deliminated by semi-colon (used last for "Contact User" if nothing else matches)
$props = @("operatingsystemversion","extensionAttribute1","extensionAttribute2","extensionAttribute3","extensionAttribute5","extensionAttribute9","extensionAttribute10","LastLogonDate")

# Get enabled computers matching $eolver or lower.
$comps = Get-ADComputer -Searchbase $searchbase -Filter {Enabled -eq $true} -Properties $props | where {$_.OperatingSystemVersion -match "\d+\.\d+ \((\d+)\)" -and $Matches.1 -ne $null -and ($Matches.1 -as [int]) -is [int] -And $Matches.1 -le $eolver}

# Filter to only systems matching "Laptop" or "Tablet", determine "Contact User", and format results.
$comps2 = $comps | where {$_.extensionAttribute5 -eq "Laptop" -Or $_.extensionAttribute5 -eq "Tablet"} | Select Name,OperatingSystemVersion,@{N="Contact User"; Expression={ if($_.extensionAttribute2 -match "([^@]+)@") { $Matches.1 } elseif (-not [string]::IsNullOrWhitespace($_.extensionAttribute10)) { $_.extensionAttribute10 -split '\\' | Select -Last 1 } else { ($_.extensionAttribute3 -split '; ' | Select -First 1) -split '\\' | Select -Last 1 } }},extensionAttribute5,LastLogonDate,extensionAttribute9,@{N="Link"; Expression={ "{0}{1}" -f $sorurl, $_.extensionAttribute1}},extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute10 | Select Name,OperatingSystemVersion,"Contact User",@{N="UserType"; Expression={$u = $null; if (-Not [string]::IsNullOrWhitespace($_.'Contact User') -and ($u = Get-ADUser $_.'Contact User' -Properties extensionAttribute2) -and -not [string]::IsNullOrEmpty($u.extensionAttribute2)) { $u.extensionAttribute2 } else { "" } }},extensionAttribute5,LastLogonDate,extensionAttribute9,"Link",extensionAttribute1,extensionAttribute2,extensionAttribute3,extensionAttribute10

# Output to file.
try {
	$comps2 | Select $csvfields | Export-CSV -NoTypeInformation $csvFilename
	Write-Host("Results for {0} computers outputed to [$csvfilename]." -f $comps2.Count)
} catch {
	Write-Error $_
}
$_ = Read-Host "Press enter to exit"

