# Creates or appends an Import.csv file using data exported from Snipe-It.
# Intended to re-import machines that have fallen out of SCCM.

# -- START CONFIG --
$asset_export_fp = '\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\SnipeIt\Exports\assets_snipeit_latest.csv'
$sccm_import_path = '\\win.ad.jhu.edu\data\sccmpack$\hsa\Import-Computers'
$sccm_import_filename = 'Import.csv'
# -- END CONFIG --

$asset_tag = Read-Host "Enter the asset tag or hit ENTER if not known"
if ([string]::IsNullOrWhitespace($asset_tag)) {
	$asset_name = Read-Host "Enter the full or partial asset name, using * for wildcard (ex: *-HPSF8S3)"
	if ([string]::IsNullOrWhitespace($asset_name)) {
		Write-Host "** Error: No name given. Exiting."
		Read-Host "Press enter to exit" | Out-Null
		exit 1
	}
}

Write-Host "** Loading assets from [$asset_export_fp], please wait..."
$assets = Import-CSV $asset_export_fp
Write-Host("** Searching [{0}] assets" -f $assets.Count)

if (-Not [string]::IsNullOrWhitespace($asset_tag)) {
	$asset = $assets | where {$_.asset_tag -eq $asset_tag}
} elseif (-Not [string]::IsNullOrWhitespace($asset_name)) {
	$asset = $assets | where {$_.name -like $asset_name}
}

$cols = @('id','asset_tag','name','assigned_to','manufacturer','model','system form factor','PC Checkboxes','SMBIOS GUID','SCCM LastActiveTime')
if ($asset.Count -gt 1 -Or [string]::IsNullOrWhitespace(($asset | Select -First 1 -ExpandProperty asset_tag))) {
	$asset | Select $cols 
	Write-Host ("** Returned [{0}] results. Make sure you have the correct asset tag / name. Otherwise, please search in Snipe-It to match the correct system." -f ($asset | Measure-Object | Select -ExpandProperty count))
	Read-Host "Press enter to exit" | Out-Null
	exit 2
} else {
	$asset = $asset | Select -First 1
	$asset | Select $cols 
	$asset_name = $asset | Select -ExpandProperty name
	$asset_guid = $asset | Select -ExpandProperty 'SMBIOS GUID'
	if ([string]::IsNullOrWhitespace($asset_name)) {
		Write-Host "** ERROR: Asset Name is blank. Aborting."
		Read-Host "Press enter to exit" | Out-Null
		exit 3
	}
	if ([string]::IsNullOrWhitespace($asset_guid)) {
		Write-Host "** ERROR: SMBIOS GUID is blank. Aborting."
		Read-Host "Press enter to exit" | Out-Null
		exit 4
	}

	$row = '"{0}","{1}",' -f $asset_name,$asset_guid

	# Save file to OneDrive Documents by default
	if (Test-Path -Path "${ENV:ONEDRIVE}\Documents") {
		$save_path = "${ENV:ONEDRIVE}\Documents"
	} else {
		$save_path = ".\"
	}

	$choice = $null
	If(Test-Path -Path "$save_path\$sccm_import_filename" -PathType Leaf) {
		Write-Host ("** Import file already found at [{0}]" -f "$save_path\$sccm_import_filename")
		$choice = Read-Host "** Append previously saved import file? (N/Y, Default: N)"
	}
	if($choice -ne "Y") {
		Clear-Content -Path "$save_path\$sccm_import_filename" | Out-Null
		Add-Content -Path "$save_path\$sccm_import_filename" -Value '"Name","SMBIOS GUID","MAC Address"'
	}
	Add-Content -Path "$save_path\$sccm_import_filename" -Value $row
	Get-Content -Path "$save_path\$sccm_import_filename"
	Write-Host "** Import file saved to [$save_path\$sccm_import_filename]"
	Write-Host "** SCCM Import Path: [$sccm_import_path]"
	$choice = $null
	$choice = Read-Host ("** Copy [$sccm_import_filename] to SCCM Import Path? (Y/N, Default: Y)")
	if ($choice -ne "N") {
		if (Copy-Item -Path "$save_path\$sccm_import_filename" -Destination "$sccm_import_path\$sccm_import_filename" -Force -PassThru) {
			Write-Host "** [$sccm_import_filename] copied to [$sccm_import_path]"
		} else {
			Write-Warning "Copied failed to [$sccm_import_path\$sccm_import_filename]"
		}
	}
}
Read-Host "Press enter to exit" | Out-Null

