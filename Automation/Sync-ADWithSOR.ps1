<#
	.SYNOPSIS
	Syncs information on assets from the System-Of-Record (Snipe-It) with AD.
	
	.DESCRIPTION
	Syncs information on assets from the System-Of-Record (Snipe-It) with AD.
	
	.NOTES
	Requirements:
	* RSAT: Active Directory PowerShell module.
	
	Created: 4-9-24
	Author: mcarras8

	Changelog
	03-14-25 - MJC - Standardized documentation and uploaded to github under ScriptsLibrary.
	04-16-24 - MJC - Added emailed error reports and changelog.
#>

# -- START CONFIGURATION --
$CSV_IMPORT_FILEPATH = "\\win.ad.jhu.edu\cloud\hsa$\ITServices\Reports\SnipeIt\Exports\assets_snipeit_latest.csv"

# Searchbases to optionally sync information from AD.
$AD_IMPORT_SEARCHBASES = @("OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu")
<#
$AD_IMPORT_SEARCHBASES = @( "OU=USS-IT,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu",
							"OU=USS-SHWB,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu",
							"OU=USS-XX,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu",
							"OU=USS-LNR,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu",
							"OU=USS-LAB,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu",
							"OU=USS-STR,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu",
							"OU=USS-DMC,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu"
)
#>

# SOR CSV column = AD attribute
# Note: CURRENT_DATE is a 
$ASSET_SYNC_MAP = @{
	Department = "department"
	serial = "serialNumber"
	asset_tag = "extensionAttribute1"
	assigned_to = "extensionAttribute2"
	"Primary Users" = "extensionAttribute3"
	location = "extensionAttribute4"
	"System Form Factor" = "extensionAttribute5"
	manufacturer = "extensionAttribute6"
	model = "extensionAttribute7"
	"PC Checkboxes" = "extensionAttribute9"
	"LastLogonUser" = "extensionAttribute10"
}
# Field to match on name.
$ASSET_FIELD_NAME = "name"
# AD attribute to enter the current date (to show last sync date).
# Only used when one of the other attributes is replaced or cleared.
$LAST_UPDATE_ATTR = "extensionAttribute8"

# Regex for grabbing the username from the assigned_to field.
$ASSET_FIELD_ASSIGNED_TO = "assigned_to"
$ASSET_REGEX_ASSIGNED_TO = "\s\(([^\s]+)\)$"
# Regex for when to include the assigned_to field in the description.
# Leave blank to always include it.
$ASSET_DESCRIPTION_REGEX_ASSIGNED_TO = "@"

# Array of columns from the SOR to use in the description (pipe-delimited).
$DESCRIPTION_FORMAT_ARRAY = @("assigned_to", "Department", "asset_tag", "location")

# Restrict the imported SOR assets based on the given scriptblock.
$ASSET_RESTRICT_WHERE_SCRIPTBLOCK = {$_.Category -eq "PC" -or $_.Category -eq "Mac"}

# Path and prefix for the Start-Transcript logfiles.
$LOGFILE_PATH = ".\Logs"
$LOGFILE_PREFIX = "sync-adwithsor"
# Maximum number of days before rotating logfile.
$LOGFILE_ROTATE_DAYS = 90

# Email configuration for reports
$EMAIL_SMTP = 'smtp.johnshopkins.edu'
# If filled out, send error reports
$EMAIL_ERROR_REPORT_FROM = 'USS IT Services <ussitservices@jhu.edu>'
# Can be string or array of strings.
$EMAIL_ERROR_REPORT_TO = @('mcarras8@jhu.edu','ussitservices@jhu.edu')

# -- END CONFIGURATION --

# -- START --

# Rotate log files
if ($LOGFILE_ROTATE_DAYS -is [int] -And $LOGFILE_ROTATE_DAYS -gt 0) {
	Get-ChildItem "${LOGFILE_PATH}\${LOGFILE_PREFIX}_*.log" | Where-Object { $_.CreationTime -lt (Get-Date).AddDays(-$LOGFILE_ROTATE_DAYS) } | Remove-Item -Force
}

# Start logging
$_logfilepath = "${LOGFILE_PATH}\${LOGFILE_PREFIX}_$(get-date -f yyyy-MM-dd).log"
Start-Transcript -Path $_logfilepath -Append

# -- START FUNCTIONS --

# Import assets from AD.
function Import-AssetsFromAD {
	param (
		[parameter(Mandatory=$true)]
		[string[]]$Properties,
		
		[parameter(Mandatory=$false)]
		[string[]]$SearchBase,
	
		[parameter(Mandatory=$false)]
		[ValidateScript({-Not [string]::IsNullOrWhitespace($_)})]
		[string]$Filter="*"
	)
	
    Write-Verbose("[Import-AssetsFromAD] Collecting all assets from AD using searchbases [{0}], this might take a while..." -f ($SearchBase -join ", "))

	$props = $Properties
	$props += @("distinguishedname")
	$props += @("description")
	
	if ($SearchBase.Count -gt 0) {
		$assets = $Searchbase | foreach { Get-ADComputer -SearchBase $_ -Filter $Filter -Properties $props}
	} else {
		$assets = Get-ADComputer -Filter $Filter -Properties $props
	}

    Write-Verbose("[{0}] [Import-AssetsFromCSV] {1} assets imported from AD" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $assets.Count)
	return $assets
}

# -- END FUNCTIONS --

$_scriptName = split-path $PSCommandPath -Leaf
$error_count = 0

if($ASSET_SYNC_MAP.ContainsKey($LAST_UPDATE_ATTR)) {
	Write-Warning("CURRENT_DATE_ATTR=[{0}] will overwrite value from ASSET_SYNC_MAP field [{1}]=[{0}]" -f $LAST_UPDATE_ATTR, $ASSET_SYNC_MAP.GetEnumerator() | where {$_.Value -eq $LAST_UPDATE_ATTR } | Select -ExpandProperty Name -First 1)
}

# Get assets from SOR imported CSV and AD. Limit to categories of PC or Mac.
$sor_assets = Import-CSV $CSV_IMPORT_FILEPATH | where $ASSET_RESTRICT_WHERE_SCRIPTBLOCK
$ad_properties = $ASSET_SYNC_MAP.Values | foreach { $_ }
$ad_assets = Import-AssetsFromAD -Properties $ad_properties -SearchBase $AD_IMPORT_SEARCHBASES -Verbose

# Loop over all matching assets.
foreach($asset in $sor_assets) {
	If (($ad_asset = $ad_assets | where {$_.Name -eq $asset.$ASSET_FIELD_NAME}) -And -Not [string]::IsNullOrEmpty($ad_asset.Name)) {
		$replace_attrs = @{}
		$clear_attrs = @()
		$assigned_to = $null
		# Set assigned_to
		if (-Not [string]::IsNullOrEmpty($ASSET_FIELD_ASSIGNED_TO) -And -Not [string]::IsNullOrEmpty($ASSET_SYNC_MAP.$ASSET_FIELD_ASSIGNED_TO)) {
			$assigned_to = $asset.$ASSET_FIELD_ASSIGNED_TO
			If(-Not [string]::IsNullOrEmpty($ASSET_REGEX_ASSIGNED_TO) -And $assigned_to -match $ASSET_REGEX_ASSIGNED_TO -And -Not [string]::IsNullOrWhitespace($Matches.1)) {
				$assigned_to = $Matches[1]
			}
		}
		# Loop over each mapped field key.
		foreach($key in $ASSET_SYNC_MAP.Keys) {
			$prop = $ASSET_SYNC_MAP.$key
			if($key -eq $ASSET_FIELD_ASSIGNED_TO -And $assigned_to -ne $null) {
				$sor_value = $assigned_to
			} else {
				$sor_value = $asset.$key
			}
			$ad_value = $ad_asset.$prop
			try {
				# For multi-value attributes like serialNumber
				if($ad_value.GetType().Name -eq "ADPropertyValueCollection") {
					$ad_value = $ad_asset.$prop | Select -First 1
				}
			} catch {}
			if($sor_value -ne $ad_value -And (-Not [string]::IsNullOrEmpty($sor_value) -Or -Not [string]::IsNullOrEmpty($ad_value))) {
				if([string]::IsNullOrEmpty($sor_value) -And -Not [string]::IsNullOrEmpty($ad_value)) {
					Write-Host("[{0}] [{1}] attr={2} [{3}] -ne [{4}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $ad_asset.Name, $prop, $ad_value, $sor_value)
					$clear_attrs += @($prop)
				} else {
					Write-Host("[{0}] [{1}] attr={2} [{3}] -ne [{4}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $ad_asset.Name, $prop, $ad_value, $sor_value)
					$replace_attrs.Add($prop, $sor_value)
				}
			}
		}
		# Loop over the fields to generate a description.
		$computed_description = @()
		foreach($key in $DESCRIPTION_FORMAT_ARRAY) {
			# Exclude the assigned_to field if matching regex
			$sor_value = $asset.$key
			if(-Not [string]::IsNullOrWhitespace($sor_value) -And ($key -ne $ASSET_FIELD_ASSIGNED_TO -Or [string]::IsNullOrEmpty($ASSET_DESCRIPTION_REGEX_ASSIGNED_TO) -Or $sor_value -match $ASSET_DESCRIPTION_REGEX_ASSIGNED_TO)) {
				$computed_description += @("{0}: {1}" -f $key,$sor_value)
			}
		}
		# Generate a description from multiple fields.
		$computed_description = $computed_description -join " | "
		if ($computed_description -ne $ad_asset.description) {
			Write-Host("[{0}] [{1}] attr=description [{2}] -ne [{3}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $ad_asset.Name, $ad_asset.description, $computed_description)
			$replace_attrs.Add("description", $computed_description)
		}
		# Replace or clear attributes for this ad object (if any).
		if($replace_attrs.Count -gt 0 -Or $clear_attrs.Count -gt 0) {
			# Add current date to replace attributes if set.
			if(-not [string]::IsNullOrWhitespace($LAST_UPDATE_ATTR)) {
				Write-Host("[{0}] [{1}] Updating last update attribute [{2}] due to changed attributes" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $ad_asset.Name, $LAST_UPDATE_ATTR)
				$replace_attrs.Add($LAST_UPDATE_ATTR, (Get-Date -Format "MM-dd-yyyy"))
			}
			if ($replace_attrs.Count -gt 0) {
				Write-Host("[{0}] [{1}] Replace Atributes: {2}" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $ad_asset.Name, ($replace_attrs.Keys -join ","))
			}
			if ($clear_attrs.Count -gt 0) {
				Write-Host("[{0}] [{1}] Clear Attributes: {2}" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $ad_asset.Name, ($clear_attrs -join ","))
			}
			if($clear_attrs.Count -gt 0) {
				try {
					Set-ADComputer $ad_asset.distinguishedname -Replace $replace_attrs -Clear $clear_attrs
				} catch {
					Write-Error $_
					$error_count++
				}
			} else {
				try {
					Set-ADComputer $ad_asset.distinguishedname -Replace $replace_attrs
				} catch {
					Write-Error $_
					$error_count++
				}
			}
		}
	}
}

Write-Host("[{0}] Caught {1} errors" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), $error_count)

# Stop logging
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

if (-Not [string]::IsNullOrWhiteSpace($EMAIL_SMTP)) {  
    # Email out notifications of any errors.
    if ($error_count -gt 0 -And -Not [string]::IsNullOrWhiteSpace($EMAIL_ERROR_REPORT_FROM) -And -Not [string]::IsNullOrWhiteSpace(($EMAIL_ERROR_REPORT_TO | Select -First 1)))	{
		$emailParams = @{
			From = $EMAIL_ERROR_REPORT_FROM
			To =  $EMAIL_ERROR_REPORT_TO
			Subject = "Errors from $_scriptName"
			Body = "There were [$error_count] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See attached logfile for more details."
			Priority = "High"
			DeliveryNotificationOption = @("OnSuccess", "OnFailure")
			SmtpServer = $EMAIL_SMTP
		}
        try {
			Send-MailMessage -Attachments $_logfilepath @emailParams
		} catch {
			Write-Error $_
			$mailParams.Body = "There were [$error_count] caught errors from [$_scriptName] running on [${ENV:COMPUTERNAME}]. See [$_logfilepath] for more details."
			Send-MailMessage @emailParams
		}
        Write-Host("[{0}] Emailed error report to [{1}]" -f ((Get-Date).toString("yyyy/MM/dd HH:mm:ss")), ($EMAIL_ERROR_REPORT_TO -join ", "))
    }
}