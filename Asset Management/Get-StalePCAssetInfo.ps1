# Quick script to cross-reference assets to be deleted with SOR (Snipe-It).
# MJC 4-10-24
$stale = Import-CSV '.\StalePCs- to be deleted.csv'
$assets = Import-CSV '\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\SnipeIt\Exports\assets_snipeit_latest.csv'
$asset_url = 'https://jh-uss.snipe-it.io/hardware/'

$stale_assets = $stale | foreach {
	$name = $_.ComputerName;
	if (($asset = $assets | where {$_.Name -eq $name -And -Not [string]::IsNullOrWhitespace($_.Name)}) -And -Not [string]::IsNullOrEmpty($asset.Name)) {
		if($asset.category -eq "Mac") {
			$checkboxes = $asset.'Apple Checkboxes'
		} else {
			$checkboxes = $asset.'PC Checkboxes'
		}
		[PSCustomObject]@{
			ComputerName = $name
			LastLogonDate = $_.LastLogonDate
			PingResult = $_.PingResult
			Action = $_.Action
			"Asset Category" = $asset.category
			"Asset Tag" = $asset.asset_tag
			"Asset Status" = $asset.status
			"Asset Checkboxes" = $checkboxes
			"Asset Assigned" = $asset.assigned_to
			"Asset Department" = $asset.Department
			"Asset Form Factor" = $asset.'System Form Factor'
			"Asset Manufacturer" = $asset.manufacturer
			"Asset Model" = $asset.model
			"Asset Link" = $asset_url + $asset.id
		}
	} else {
		# Item not found
		[PSCustomObject]@{
			ComputerName = $name
			LastLogonDate = $_.LastLogonDate
			PingResult = $_.PingResult
			Action = $_.Action
			"Asset Category" = ""
			"Asset Tag" = ""
			"Asset Status" = ""
			"Asset Checkboxes" = ""
			"Asset Assigned" = ""
			"Asset Department" = ""
			"Asset Form Factor" = ""
			"Asset Manufacturer" = ""
			"Asset Model" = ""
			"Asset Link" = ""
		}
	}
}
$stale_assets | Export-CSV -NoTypeInformation '.\StalePCs- to be deleted with asset info.csv'
