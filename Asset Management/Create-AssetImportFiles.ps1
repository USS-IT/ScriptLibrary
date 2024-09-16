# MJC 5-8-24
# Creates import scripts for Snipe-It, SCCM, and JHARS from a Dell report (exported as CSV).
# Copies imported sheets to $IMPORTSHEET_COPYTO_FILEPATH
param(
	# The import sheet as a CSV file.
	[Parameter(Position=0,Mandatory=$true)]
	[string] $CSV,
	
	# Assume defaults for interactive prompts.
	[switch] $NoPrompts
)

## -- START CONFIGURATION --

# Filepath to copy successful import CSV.
$IMPORTSHEET_COPYTO_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\Dell\Import_Sheets"
$SCCM_IMPORT_COPYTO_PATH = "\\win.ad.jhu.edu\data\sccmpack$\hsa\Import-Computers"
$SNIPEIT_IMPORT_COPYTO_PATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\SnipeIt"

# SnipeIt headers
#Serial,Name,Manufacturer,Model,Purchasing date,Warranty Exp.,End of Life,Location,Category,System Form Factor,PC Checkboxes,SMBIOS GUID,Supplier,Order Number,Purchase Date

# Prefixes for generating names from serial #s.
$NAME_PREFIX = "USS-XX-"

# All mac address columns.
$MAC_COLUMNS = "Pass Through Mac", "Embedded MAC Address 1"

# JHARS Bulk Import file info
$JHARS_IMPORT_FILEPATH = "JHARS_Bulk_Import.csv"
$JHARS_MAC_COLUMNS = $MAC_COLUMNS

# SCCM Import file info
# Mapping supports multiple possible mappings as array value.
# If an array, the first value found from a valid column will be used.
# e.g. "LOM MAC Address" = "Embedded MAC Address 1", "MAC Address"
$SCCM_IMPORT_FILEPATH = "Import.csv"
$SCCM_HEADER = "Name", "SMBIOS GUID", "MAC Address"
$SCCM_HEADER_MAP = @{
	"SMBIOS GUID" = "UUID"
	"MAC Address" = "Embedded MAC Address 1"
	"Fallback MAC Address" = "Pass Through MAC"
}

# Snipe-It Import file info
$SNIPEIT_IMPORT_FILEPATH = "Snipeit-Import.csv"

# Not all columns are required.
# TODO: Find out why creating new Suppliers is erroring.
#$SNIPEIT_HEADER = "Serial","Name","Manufacturer","Model","Purchasing date","Warranty Exp.","End of Life","Location","Category","System Form Factor","PC Checkboxes","SMBIOS GUID","Supplier","Order Number","LOM MAC Address","Pass-Through MAC Address"
$SNIPEIT_HEADER = "Serial","Name","Manufacturer","Model","Purchasing date","Warranty Exp.","End of Life","Location","Category","System Form Factor","PC Checkboxes","SMBIOS GUID","Order Number","LOM MAC Address","Pass-Through MAC Address"

# Mapping supports multiple possible mappings as array value.
# If an array, the first value found from a valid column will be used.
# e.g. "LOM MAC Address" = "Embedded MAC Address 1", "MAC Address"
$SNIPEIT_HEADER_MAP = @{
	"Order Number" = "Order Number"
	"Serial" = "Dell Service Tag"
	"Model" = "Chassis Description"
	"System Form Factor" = "Chassis Style"
	"Purchasing Date" = "Order Date"
	"SMBIOS GUID" = "UUID"
	#"Supplier" = "Company Name"
	"LOM MAC Address" = "Embedded MAC Address 1"
	"Pass-Through MAC Address" = "Pass Through MAC"
}

# Default values used if not mapped or mapping is blank.
$SNIPEIT_DEFAULT_VALUES = @{
	"PC Checkboxes" = "Verified"
	"Location" = "MD,HW,Krieger Hall,Floor1,Rm160,Storage"
	"Category" = "PC"
	"Manufacturer" = "Dell, Inc."
}

$WARRANTY_LENGTH_YEARS = 3
$EOL_LENGTH_YEARS = $WARRANTY_LENGTH_YEARS + 1

$CHASSIS_STYLE_MAP = @{
	"NOTEBOOK" = "Laptop"
	"SFF" = "Desktop"
}

## -- END CONFIGURATION --

if (-Not (Test-Path $CSV -PathType Leaf)) {
	Write-Error "Invalid or unknown CSV filepath [$CSV]"
	return -1
}

$csvImport = Import-CSV $CSV
Write-Host('[{0}] Loaded [{1}] assets from CSV file [{2}]...' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"),$csvImport.Count,$CSV) 

function Format-GUID($guid) {
	# Example: 4C4C4544-0050-5310-8046-B1C04F385333
	return "{0}-{1}-{2}-{3}-{4}" -f $guid.SubString(0,8), $guid.SubString(8,4), $guid.SubString(12,4), $guid.SubString(16,4), $guid.SubString(20,12)
}

function Format-MAC($mac) {
	return ($mac -split '(..)' -ne '') -join ':'
}

# Only continue if the first entry is valid.
if($csvImport -And -Not [string]::IsNullOrEmpty(($csvImport | Select -First 1 -ExpandProperty $SNIPEIT_HEADER_MAP.Serial))) {
	$snipeit_import = $csvImport | foreach {
		$row = $_
		$o = [PSCustomObject]@{}
		foreach($ecol in $SNIPEIT_HEADER) {
			$val = $null
			if($ecol -eq "Name") {
				$val = $NAME_PREFIX + $row.($SNIPEIT_HEADER_MAP["Serial"])
			} elseif ($ecol -eq "Warranty Exp.") {
				$date = $row.($SNIPEIT_HEADER_MAP["Purchasing Date"]) -as [DateTime]
				if ($date -is [DateTime]) {
					$val = $date.AddYears($WARRANTY_LENGTH_YEARS).ToString("MM/dd/yyyy")
				}
			} elseif ($ecol -eq "End of Life") {
				$date = $row.($SNIPEIT_HEADER_MAP["Purchasing Date"]) -as [DateTime]
				if ($date -is [DateTime]) {
					$val = $date.AddYears($EOL_LENGTH_YEARS).ToString("MM/dd/yyyy")
				}
			} else {
				# Check header map
				$icol = $SNIPEIT_HEADER_MAP[$ecol]
				if ($icol -is [array]) {
					foreach($col in $icol) {
						$val = $row.$col
						if (-Not [string]::IsNullOrEmpty($val)) {
							break
						}
					}
				}
				if (-Not [string]::IsNullOrEmpty($icol)) {
					$val = $row.$icol
					# Map other values
					if(-Not [string]::IsNullOrEmpty($val)) {
						if ($ecol -eq "System Form Factor") {
							$ff = $CHASSIS_STYLE_MAP[$val]
							if(-Not [string]::IsNullOrEmpty($ff)) {
								$val = $ff
							}
						} elseif ($ecol -eq "SMBIOS GUID" -And $val.Length -eq 32) {
							$val = Format-GUID $val
						} elseif ($icol -in $MAC_COLUMNS -And $val -ne 'NONE') {
							$val = Format-MAC $val
						}
					}
				} else {
					$val = $SNIPEIT_DEFAULT_VALUES[$ecol]
				}
			}
			Add-Member -InputObject $o -MemberType NoteProperty -Name $ecol -Value $val -Force
		}
		$o
	}
	$snipeit_import | Export-CSV -NoTypeInformation -Force $SNIPEIT_IMPORT_FILEPATH
	
	If (Test-Path $SNIPEIT_IMPORT_FILEPATH -PathType Leaf) {
		Write-Host('[{0}] Created Snipe-It import file [{1}]' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $SNIPEIT_IMPORT_FILEPATH)
		Get-Content $SNIPEIT_IMPORT_FILEPATH -TotalCount 2
		Write-Host "..."
		Write-Host('[{0}] Snipe-It import path: {1}' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $SNIPEIT_IMPORT_COPYTO_PATH)
		$choice = $null
		if (-Not $NoPrompts) {
			$choice = Read-Host ">> Copy Snipe-It import file to designated path? (Y/N)"
		}
		if ($choice -ne "N") {
			if (Copy-Item -Path $SNIPEIT_IMPORT_FILEPATH -Destination $SNIPEIT_IMPORT_COPYTO_PATH -Force -Passthru) {
				Write-Host('[{0}] Copied [{1}] to [{2}]' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $SNIPEIT_IMPORT_FILEPATH, $SNIPEIT_IMPORT_COPYTO_PATH)
			}
		}
	}
	
	$sccm_import = $csvImport | foreach {
		$row = $_
		$o = [PSCustomObject]@{}
		foreach($ecol in $SCCM_HEADER) {
			$val = $null
			if($ecol -eq "Name") {
				$val = $NAME_PREFIX + $row.($SNIPEIT_HEADER_MAP["Serial"])
			} else {
				# Check header map
				$icol = $SCCM_HEADER_MAP[$ecol]
				if ($icol -is [array]) {
					foreach($col in $icol) {
						$val = $row.$col
						if (-Not [string]::IsNullOrEmpty($val)) {
							break
						}
					}
				}
				if (-Not [string]::IsNullOrEmpty($icol)) {
					$val = $row.$icol
					# Map other values
					if(-Not [string]::IsNullOrEmpty($val)) {
						if ($icol -in $MAC_COLUMNS -And $val -ne 'NONE') {
							$val = Format-MAC $val
						} elseif ($ecol -eq "SMBIOS GUID" -And $val.Length -eq 32) {
							$val = Format-GUID $val
						}
					# Fallback to passthrough MAC if embedded MAC is blank or doesn't exist
					} elseif ($icol -eq "MAC Address") {
						$val = $row."Fallback MAC Address"
						if (-Not [string]::IsNullOrEmpty($val)) {
							$val = Format-MAC $val
						}
					}
				}
			}
			Add-Member -InputObject $o -MemberType NoteProperty -Name $ecol -Value $val -Force
		}
		$o
	}
	$sccm_import | Export-CSV -NoTypeInformation -Force $SCCM_IMPORT_FILEPATH
	
	If (Test-Path $SCCM_IMPORT_FILEPATH -PathType Leaf) {
		Write-Host('[{0}] Created SCCM import file [{1}]' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $SCCM_IMPORT_FILEPATH)
		Get-Content $SCCM_IMPORT_FILEPATH -TotalCount 2
		Write-Host "..."
		Write-Host('[{0}] SCCM import path: {1}' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $SCCM_IMPORT_COPYTO_PATH)
		$choice = $null
		if (-Not $NoPrompts) {
			$choice = Read-Host ">> Copy SCCM import file to designated path? (Y/N)"
		}
		if ($choice -ne "N") {
			if (Copy-Item -Path $SCCM_IMPORT_FILEPATH -Destination $SCCM_IMPORT_COPYTO_PATH -Force -Passthru) {
				Write-Host('[{0}] Copied [{1}] to [{2}]' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $SCCM_IMPORT_FILEPATH, $SCCM_IMPORT_COPYTO_PATH)
			}
		}
	}
	
	if($csvImport -And $csvImport.Count -gt 0) {
		Clear-Content -Path $JHARS_IMPORT_FILEPATH
		$csvImport | foreach {
			$row = $_
			foreach($ecol in $JHARS_MAC_COLUMNS) {
				$val = $row.$ecol
				if (-Not [string]::IsNullOrEmpty($val) -And $val -ne 'NONE') {
					$val = Format-MAC $val
					if (-not [string]::IsNullOrEmpty($val)) {
						Add-Content -Path $JHARS_IMPORT_FILEPATH -Value $val
					}
				}
			}
		}
		
		Write-Host('[{0}] Created JHARS import file [{1}]' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $JHARS_IMPORT_FILEPATH)
		Write-Host('>> Please import the JHARS info manually using the bulk upload tool.')
	}
	
	if (Copy-Item -Path $CSV -Destination $IMPORTSHEET_COPYTO_PATH -Force -Passthru) {
		Write-Host('[{0}] Copied import sheet [{2}] to [{1}]' -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $IMPORTSHEET_COPYTO_PATH, $CSV)
	}
}