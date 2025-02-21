# Returns previous import files matching pattern within the last 90 days.
$name = Read-Host "Enter or paste name"
# How far back to search.
$days = 90
# Path to files with wildcard.
$path = "\\win.ad.jhu.edu\data\sccmpack$\uss\Import-Computers\Imported\*.csv"
$results = gci $path | where {$_.CreationTime -ge (Get-Date).AddDays(-$days) -And (Select-String -Path $_.FullName -Pattern $name -SimpleMatch -Quiet)}
if(-Not $results -And $results.Count -eq 0) {
	Write-Host "[$name] not found in [$path]."
} else {
	$results
}
Read-Host "Press enter to exit" | Out-Null
