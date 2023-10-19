# Output a user's groups to a CSV file or show it in a pop-up.
# MJC 6-9-23
$username = Read-Host "Enter JHED"
$outputFile = Read-Host "Output CSV file (leave blank to show in pop-up)"
$ussOnly = Read-Host "USS OU Groups Only (Y/N, Default: Y)"

$results = Get-ADUser $username -Properties memberof | Select -ExpandProperty memberof
if ($ussOnly -ne "N") {
	$results = $results | where {$_ -like "*,OU=USS,*"}
}
$results = $results | foreach { if ($_ -match "CN=([^,]+),") { $matches[1] } }

if (-Not [string]::IsNullOrEmpty($outputFile)) {
	$results | Export-CSV -NoTypeInformation $outputFile
	Write-Host "Exported to [$outputFile]"
} else {
	$results
	$results | Out-GridView -Title "AD Groups for $username"
}
Read-Host "Press enter to exit" | Out-Null