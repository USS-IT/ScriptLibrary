# Output the members of a group to a CSV file or show it in a pop-up.
# MJC 6-8-23
$groupname = Read-Host "Enter AD group name"
$outputFile = Read-Host "Output to CSV file (leave blank to show in pop-up)"
try {
	$results = Get-ADGroupMember -Recursive $groupname | foreach { Get-ADUser $_ -Properties Department,Company,DisplayName,mail,distinguishedname,extensionattribute2 } | Select Name,mail,DisplayName,Department,Company,extensionattribute2,distinguishedname
} catch {
	$results = Get-ADGroup $groupname | Select -Expandproperty member | foreach { Get-ADUser $_ -Properties Department,Company,DisplayName,mail,distinguishedname,extensionattribute2 }
}
$results = $results | Select Name,mail,DisplayName,Department,Company,extensionattribute2,distinguishedname

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

