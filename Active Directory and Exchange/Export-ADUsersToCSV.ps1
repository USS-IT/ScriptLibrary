<#
    .SYNOPSIS
    Output the members of a group to CSV file or show it in a pop-up.
    
	.DESCRIPTION
    Output the members of a group to CSV file or show it in a pop-up. This includes nested members.
	
    .NOTES
	Created: 6-8-23
    	Author: mcarras8
#>
$groupname = Read-Host "Enter AD group name"
$outputFile = Read-Host "Output to CSV file (leave blank to show in pop-up)"
try {
	$results = Get-ADGroupMember -Recursive $groupname
	$userResults = $results | where {$_.objectClass -eq "user"} | foreach { Get-ADUser $_ -Properties Department,Company,DisplayName,mail,distinguishedname,extensionattribute2 } | Select Name,mail,DisplayName,Department,Company,extensionattribute2,distinguishedname
	if ($userResults -ne $null -And $userResults -isnot [array]) {
		$userResults = @($userResults)
	}
	$otherResults = $results | where {$_.objectClass -ne "user"} | Select Name,@{N="extensionattribute2"; Expression={$_.objectClass}},distinguishedname
	if ($otherResults -ne $null -And $otherResults -isnot [array]) {
		$otherResults = @($otherResults)
	}
	$results = $userResults + $otherResults
} catch {
	$results = Get-ADGroup $groupname -Properties member | Select -Expandproperty member | foreach { Get-ADUser $_ -Properties Department,Company,DisplayName,mail,distinguishedname,extensionattribute2 } | Select Name,mail,DisplayName,Department,Company,extensionattribute2,distinguishedname
}

if (-Not [string]::IsNullOrEmpty($outputFile)) {
	if($outputFile -notlike "*\*") {
		$outputFile = "{0}\{1}" -f ${ENV:OneDrive}, $outputFile
	}
	$results | Export-CSV -NoTypeInformation $outputFile
	Write-Host "Exported to [$outputFile]"
} else {
	$results | Out-GridView
}
Read-Host "Press enter to exit" | Out-Null

