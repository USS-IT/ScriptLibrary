# Output a user's groups to a CSV file or show it in a pop-up.
# MJC 6-9-23
$username = Read-Host "Enter JHED"
$outputFile = Read-Host "Output CSV file (leave blank to show in pop-up)"
$ussOnly = Read-Host "USS OU Groups Only (Y/N, Default: Y)"
$getRecursive = Read-Host "Get all parent groups (may take longer) (Y/N, Default: Y)"

Write-Host "Fetching user..."
$user = Get-ADUser $username -Properties distinguishedname,memberof
if (-Not [string]::IsNullOrEmpty($user.distinguishedname)) {
	Write-Host "Fetching groups, please wait..."
	if ($getRecursive -ne "N") {
		# Use tokenGroups to get all recursive groups
		$results = ($user | Get-ADUser -Properties tokenGroups).tokenGroups | Get-ADGroup
		if ($ussOnly -ne "N") {
			$results = $results | where {$_.distinguishedname -like "*,OU=USS,*"}
		}
		$results = $results | Select Name,@{N="Parent Group"; Expression={ if($_ -in $user.memberof) { "" } else { "Y" } }},distinguishedname | Sort-Object -Property Name
	} else {
		$results = $user | Select -ExpandProperty memberof
		if ($ussOnly -ne "N") {
			$results = $results | where {$_ -like "*,OU=USS,*"}
		}
		$results = $results | Select @{N="Name"; Expression= if ($_ -match "CN=([^,]+),") { $matches[1] } else { "" }},@{N="distinguishedname"; Expression={$_}} | Sort-Object -Property Name
	}
	
	if (-Not [string]::IsNullOrEmpty($outputFile)) {
		if($outputFile -notlike "*\*") {
			$outputFile = "{0}\{1}" -f ${ENV:OneDrive}, $outputFile
		}
		$results | Export-CSV -NoTypeInformation $outputFile
		Write-Host "Exported to [$outputFile]"
	} else {
		$results
		$results | Out-GridView -Title "AD Groups for $username"
	}
}
Read-Host "Press enter to exit" | Out-Null