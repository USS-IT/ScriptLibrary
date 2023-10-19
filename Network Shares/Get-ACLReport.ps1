# Returns changed security permissions for folders with broken inheritence.
# Also includes all permissions on root.
# This may take a while to complete when run on a remote network share.
# MJC 2-3-23
$path = Read-Host "Enter or paste full path"
$doAllDir = Read-Host "Include all directories, not just ones with broken inheritance? (N/Y)"
$outputFile = Read-Host "Output to CSV file (leave blank to show in pop-up)"
# Include root folder
$results = (@(Get-Item $path) + @(Get-ChildItem $path -Directory -Recurse)) | ForEach-Object { 
	$folder = $_
	Get-ACL -Path $folder.FullName | ForEach-Object { 
		$owner = $_.Owner
		foreach ($access in ($_.Access|where {$doAllDir -eq "Y" -OR $_.IsInherited -eq $False -OR $folder.FullName -eq $path})) { 
			[PSCustomObject]@{
				'Path' = $folder.FullName
				'Owner' = $owner
				'AD Group or User' = $access.IdentityReference
				'Permissions' = $access.FileSystemRights
				'Inherited' = $access.IsInherited
			}
		} 
	} 
}
if (-Not [string]::IsNullOrEmpty($outputFile)) {
	$results | Export-CSV -NoTypeInformation $outputFile
	Write-Host "Exported to [$outputFile]"
} else {
	$results | Out-GridView
}
Read-Host "Press enter to exit" | Out-Null

