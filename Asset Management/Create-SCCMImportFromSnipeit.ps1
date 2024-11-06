# Creates or appends an Import.csv file using data exported from Snipe-It.
# Intended to re-import machines that have fallen out of SCCM.

# -- START CONFIG --
# Where to look for exports from Snipe-It.
$asset_export_fp = '\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\SnipeIt\Exports\assets_snipeit_latest.csv'
# What to name the import file.
$sccm_import_filename = 'Import.csv'
# Where to copy the import file.
$sccm_import_path = '\\win.ad.jhu.edu\data\sccmpack$\hsa\Import-Computers'
# -- END CONFIG --

$assetCols = @('id','asset_tag','name','assigned_to','manufacturer','model','system form factor','PC Checkboxes','SMBIOS GUID','SCCM LastActiveTime')
$assets = $null
$outputRows = @()
$getAssetChoice = "Y"
$loopCount = 0
while (-Not [string]::IsNullOrWhitespace($getAssetChoice) -And $getAssetChoice -ne "N") {
	$asset_tag = $null
	$asset_name = $null
	$asset_tag = Read-Host "Enter the asset tag or hit ENTER if not known"
	if ([string]::IsNullOrWhitespace($asset_tag)) {
		$asset_name = Read-Host "Enter the full or partial asset name, using * for wildcard (ex: *-HPSF8S3)"
		if ([string]::IsNullOrWhitespace($asset_name)) {
			Write-Host "** Error: No name or asset tag given."
			# Exit out if first entry
			if ($loopCount -eq 0) {
				Read-Host "Press enter to exit" | Out-Null
				exit 1
			}
		}
	}

	if($assets -eq $null) {
		Write-Host "** Loading assets from [$asset_export_fp], please wait..."
		# Check last modified time on export in case exporting has stopped.
		$fileinfo = Get-Item $asset_export_fp
		if ($fileinfo.LastWriteTime -is [DateTime] -And $fileinfo.LastWriteTime -lt (Get-Date).AddDays(-7)) {
			Write-Warning "Asset export from Snipe-It is over 7 days old. Information may be out of date."
		}
		# Import file
		$assets = Import-CSV $asset_export_fp
		if ([string]::IsNullOrWhitespace(($assets | Select -First 1 -ExpandProperty asset_tag))) {
			Write-Host "** ERROR: No valid assets found in export file. Aborting."
			Read-Host "Press enter to exit" | Out-Null
			exit 2
		}
	}

	# Search for assets from file.
	Write-Host("** Searching [{0}] assets" -f $assets.Count)
	if (-Not [string]::IsNullOrWhitespace($asset_tag)) {
		$asset = $assets | where {$_.asset_tag -eq $asset_tag}
	} elseif (-Not [string]::IsNullOrWhitespace($asset_name)) {
		$asset = $assets | where {$_.name -like $asset_name}
	}

	if ($asset.Count -gt 1 -Or [string]::IsNullOrWhitespace(($asset | Select -First 1 -ExpandProperty asset_tag))) {
		$asset | Select $assetCols
		Write-Host ("** Returned [{0}] results. Make sure you have the correct asset tag / name. Otherwise, please search in Snipe-It to match the correct system." -f ($asset | Measure-Object | Select -ExpandProperty count))
	} else {
		$asset_name = $null
		$asset_guid = $null
		$asset = $asset | Select -First 1
		$asset | Select $assetCols
		$asset_name = $asset | Select -ExpandProperty name
		$asset_guid = $asset | Select -ExpandProperty 'SMBIOS GUID'
		$asset_pccheckboxes = $asset | Select -ExpandProperty 'PC Checkboxes'
		if ([string]::IsNullOrWhitespace($asset_name)) {
			Write-Host "** ERROR: Invalid entry. Asset Name is blank."
		} elseif ([string]::IsNullOrWhitespace($asset_guid)) {
			Write-Host "** ERROR: Invalid entry. SMBIOS GUID is blank."
		} else {
			# Check if asset already exists in SCCM.
			if ($asset_pccheckboxes -match "Exists in SCCM") {
				Write-Warning "[$asset_name] already appears to exist in SCCM. Try PXE Boot or double-check SCCM."
			}
			# Search AD for computer if RSAT tools are installed.
			try {
				if (($adcomp = Get-ADComputer $asset_name -ErrorAction SilentlyContinue) -And -Not [string]::IsNullOrEmpty($adcomp.Name) -And -Not $adcomp.Enabled) {
					Write-Warning ("[$asset_name] is currently disabled in AD. DN={0}" -f $adcomp.DistinguishedName)
				}
			} catch {
			}
			
			$outputRows += @('"{0}","{1}",' -f $asset_name,$asset_guid)
			
			$getAssetChoice = Read-Host "Search for another asset? (N/Y, Default: N)"
		}
	}
	
	$loopCount++
}
if ($outputRows.Count -gt 0) {
	# Save file to OneDrive Documents by default
	if (Test-Path -Path "${ENV:ONEDRIVE}\Documents") {
		$save_path = "${ENV:ONEDRIVE}\Documents"
	} else {
		$save_path = ".\"
	}

	$choice = $null
	$save_fp = "$save_path\$sccm_import_filename"
	If(Test-Path -Path $save_fp -PathType Leaf) {
		Write-Host ("** Import file already found at [{0}]. Select option to append or replace this file." -f $save_fp)
		$choice = Read-Host "** Append previously saved import file? (N/Y, Default: N)"
	}
	if($choice -ne "Y") {
		Clear-Content -Path $save_fp | Out-Null
		Add-Content -Path $save_fp -Value '"Name","SMBIOS GUID","MAC Address"'
	}
	foreach ($row in $outputRows) {
		Add-Content -Path $save_fp -Value $row
	}
	Get-Content -Path $save_fp
	Write-Host "** Import file saved to [$save_fp]"
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

