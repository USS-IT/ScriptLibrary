<#
	.SYNOPSIS
	Checks or installs latest Dell BIOS for this system.
	
	.DESCRIPTION
	Checks or installs latest Dell BIOS for this system.
	
	.PARAMETER Check
	Return 0 for no update or 1 for an update. Return values below zero on error. This is the default if no parameters are given.
	
	.PARAMETER Install
	Installs the latest BIOS if it's an update.

	.PARAMETER Download
	Downloads the BIOS update if it's an update.
	
	.PARAMETER DownloadSKU
	Downloads the latest available BIOS update for the given SKU.
	
	.PARAMETER DownloadCSV
	Downloads latest available BIOS updates for SKUs in the given CSV file with headers "SKU","Path". If Path is empty, uses TempDir.
	
	.PARAMETER SkipPowerCheck
	Skips the AC Power check when installing BIOSes.
	
	.PARAMETER BIOSPassword
	BIOS password for the install package, if needed
	
	.PARAMETER Path
	Optional directory path to download the BIOS update if -Download or -DownloadSKU is given. Defaults to TempDir.
	
	.PARAMETER TempDir
	Directory to use for temporary files. Will create directory if not found. Default is $env:TEMP\Get-DellBIOSUpdate
	
	.PARAMETER LogDir
	Directory to store / look for log files. Will create directory if not found. Defaults to C:\Dell\Logs
	
	.PARAMETER LogLevel
	Level of console logging.
	
	.OUTPUTS
	Exit code 0 - Check: No BIOS update available. Download: Successful. Install: update was successful but no restart is required (should not happen).
	Exit code 1 - Check: BIOS update available. Download: Unsuccessful. Install: update was unsuccessful.
	Exit code 3010 - Install: BIOS update successful, restart required
	Exit code -1 - Install: package install was unsuccessful.
	Exit code -2 - Install: Not a Dell system.
	Exit code -3 - Install: No AC power detected, and -SkipPowerCheck not given.
	Exit code -4 - DownloadCSV: Error importing CSV.
	Exit code -9 - Function thrown error. For -DownloadCSV, this may include any one of the download attempts.
	Exit code -10 - Uncaught or unexpected error
	All other exit codes - Install: Returned from BIOS Dell Update Package (see notes)
	
	.EXAMPLE
	Get-DellBIOSUpdate.ps1 -Install
	
	.NOTES
	Author: Matt Carras (mcarras8) 6-17-26
	
	Some code copied from https://github.com/gwblok/garytown/blob/master/hardware/Dell/CommandUpdate/EMPS/Dell-EMPS.ps1
	
	In task sequences, you can check for exit codes using _SMSTSLastActionRetCode immediately after calling the PowerShell script and adding 1 to the success values.
	
	Use -DownloadCSV to download updates for multiple known Dell SKUs. You can lookup a system's SKU with (Get-CimInstance -ClassName Win32_ComputerSystem).SystemSKUNumber.
	The CSV should have headers "SKU","Path", where Path is an optional download directory path. If Path is empty, it defaults to TempDir.
	
	Dell BIOS Update Package Exit Code
	https://www.dell.com/support/kbdoc/en-us/000148745/dup-bios-updates
		
	-1	Dell Command Update code	Unsuccessful
		DCU terminating the BIOS execution due to timeout
	0	SUCCESSFUL	Success
		The update was successful.
	1	UNSUCCESSFUL (FAILURE)	Unsuccessful
		An error occurred during the update process; the update was not successful.
	2	REBOOT_REQUIRED	Reboot required	
		You must restart the computer to apply the updates.
	3	DEP_SOFT_ERROR	Soft dependency error	
		Some possible explanations are:
		You attempted to update to the same version of the software.
		You tried to downgrade to a previous version of the software.
		To avoid receiving this error, provide the /f option.
	4	DEP_HARD_ERROR	Hard dependency error
		The required prerequisite software was not found on your computer. The update was unsuccessful because the computer did not meet BIOS, driver, or firmware prerequisites for the update to be applied, or because no supported device was found on the target computer. The DUP enforces this check and blocks an update from being applied if the prerequisite is not met, preventing the computer from reaching an invalid configuration state. The prerequisite can be met by applying for another DUP, if available. In this case, the other package should be applied before the current one so that both updates can succeed. A DEP_HARD_ERROR cannot be suppressed by using the /f switch.

		The DUP is not applicable to the computer. Some possible explanations are:
		The DUP does not support the operating system.
		The computer does not support the DUP.
	5	QUAL_HARD_ERROR	Qualification error
		A QUAL_HARD_ERROR cannot be suppressed by using the /f switch.
	6	REBOOTING_SYSTEM	Rebooting computer
		The computer is being rebooted.
	7	Password validation error	Unsuccessful
		Password not provided or incorrect password provided for BIOS execution
	8	DOWNGRADE_BAN	Requested Downgrade is not allowed.
		Downgrading the BIOS to the version run is not allowed.
	9	RPM_VERIFY_FAILED	RPM verification has failed
		The Linux DUP framework uses RPM verification to ensure the security of all DUP-dependent Linux utilities. If security is compromised, the framework displays a message and an RPM Verify Legend, and then exits with exit code 9.
	10	EC_UNSPECIFIED_ERROR	Some other error
		This exit code is for all errors that have not been specified in BIOS exit codes 0-9. That is, battery error, EC error, HW failure, so forth.
#>

[CmdletBinding(DefaultParameterSetName = 'Check')]
param( 
	[Parameter(ParameterSetName = 'Check')]
	[switch] $Check,
	
	[Parameter(ParameterSetName = 'Install')]
	[switch] $Install,
	
	[Parameter(ParameterSetName = 'Download')]
	[switch] $Download,
	
	[Parameter(ParameterSetName = 'DownloadSKU')]
	[string] $DownloadSKU,
	
	[Parameter(ParameterSetName = 'DownloadCSV')]
	[string] $DownloadCSV,
	
	[Parameter(ParameterSetName = 'Install')]
	[switch] $SkipPowerCheck,
	
	[Parameter(ParameterSetName = 'Install')]
	[string] $BIOSPassword,
	
	[Parameter(ParameterSetName = 'Download')]
	[Parameter(ParameterSetName = 'DownloadSKU')]
	[string] $Path,
	
	[Parameter(ParameterSetName = 'Install')]
	[Parameter(ParameterSetName = 'Download')]
	[Parameter(ParameterSetName = 'Check')]
	[Parameter(ParameterSetName = 'DownloadSKU')]
	[Parameter(ParameterSetName = 'DownloadCSV')]
	[string] $TempDir,
	
	[Parameter(ParameterSetName = 'Install')]
	[Parameter(ParameterSetName = 'Download')]
	[Parameter(ParameterSetName = 'Check')]
	[Parameter(ParameterSetName = 'DownloadSKU')]
	[Parameter(ParameterSetName = 'DownloadCSV')]
	[AllowEmptyString()]
	[string] $LogDir = "C:\Dell\Logs",
	
	[Parameter(ParameterSetName = 'Install')]
	[Parameter(ParameterSetName = 'Download')]
	[Parameter(ParameterSetName = 'Check')]
	[Parameter(ParameterSetName = 'DownloadSKU')]
	[Parameter(ParameterSetName = 'DownloadCSV')]
	[int] $LogLevel = 0
)

# -- START FUNCTIONS --
function Invoke-WebRequestDownload {
	<#
		.SYNOPSIS
		Downloads a file with error-handling and retries
		
		.DESCRIPTION
		Downloads a file with error-handling and retries
		
		.PARAMETER Uri
		Uri path to web resource to download
		
		.PARAMETER OutFile
		Path for output file
		
		.PARAMETER MaxRetries
		Maximum number of retries on supported errors. Default is 5.
		
		.PARAMETER RetryDelaySeconds
		Number of seconds to sleep between attempts. Default is 3.
		
		.PARAMETER UseExponentialBackoff
		If set, increase delay exponentially based on number of attempts.
		
		.OUTPUTS
		Path to output file if successful
		
		.NOTES
		Retries on invalid status codes, 408, 429, or >= 500
	#>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter(Mandatory)]
        [string]$OutFile,
		
		[Parameter(Mandatory=$false)]
		[ValidateRange(0, [int]::MaxValue)]
        [int]$MaxRetries = 5,
		
		[Parameter(Mandatory=$false)]
		[ValidateRange(0, [int]::MaxValue)]
        [int]$RetryDelaySeconds = 3,

		[Parameter(Mandatory=$false)]
        [switch]$UseExponentialBackoff
    )

    $attempt = 0

    while ($attempt -lt $MaxRetries -Or $MaxRetries -eq 0) {
        try {
            $attempt++

            Write-Verbose "Attempt $attempt - Downloading [$Uri]"

            Invoke-WebRequest -Uri $Uri -OutFile $OutFile -UseBasicParsing -ErrorAction Stop

			if (!(Test-Path $OutFile)) { 
				Write-Error "[$OutFile] not found after download"
				throw
			}
			
            # Success
            Write-Verbose "Download to [$OutFile] succeeded"
			
			return
        }
        catch {
            $isLastAttempt = ($attempt -ge $MaxRetries)

            # Extract useful error info
            $message = $_.Exception.Message

            Write-Warning "Attempt $attempt failed: $message"

            # Optional: only retry transient HTTP errors
            $statusCode = $null
            if ($_.Exception.Response -and $_.Exception.Response.StatusCode) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            $retryable =
                -not $statusCode -or               # network/DNS/etc
                ($statusCode -ge 500) -or         # server errors
                ($statusCode -in 408,429)         # timeout / throttling

            if (-not $retryable -or $isLastAttempt) {
                break
            }

            # Delay logic
            if ($UseExponentialBackoff) {
                $delay = [math]::Pow($RetryDelaySeconds, $attempts)
            }
            else {
                $delay = $RetryDelaySeconds
            }

            Write-Verbose "Retrying in $delay seconds..."
            Start-Sleep -Seconds $delay
        }
    }
	
	Write-Error "Download failed after $attempt attempts."
    throw
}

function Get-DellSupportedModels {
	<#
		.SYNOPSIS
		Returns array of supported models from Dell's CatalogIndexPC.cab
		
		.DESCRIPTION
		Returns array of supported models from Dell's CatalogIndexPC.cab. Can return cached results in same session.
	
		.OUTPUTS
		Supported models with SystemID, Model, URL, and Date
		
		.NOTES
		SystemID = System SKU.
		URL = Path on http://downloads.dell.com/ to CatalogIndexPCModel.xml. This contains all the driver and software downloads for that SKU-specific Model.
	#>
	[CmdletBinding()]
	param()
	
	# First check cache.
	$SupportedModelsObject = $script:_CachedSupportedModels
	if (($SupportedModelsObject | Measure).Count -gt 0) {
		Write-Verbose "Using cached supported models"
	} else {
		$temproot = $script:_TEMPDIR
		if (!(Test-Path $script:_TEMPDIR)) { $null = New-Item -Path $script:_TEMPDIR -ItemType Directory -Force }
		
		$CabPathIndex = "$temproot\DellCabDownloads\CatalogIndexPC.cab"
		$DellCabExtractPath = "$temproot\DellCabDownloads\DellCabExtract"
		if (!(Test-Path $DellCabExtractPath)){$null = New-Item -Path $DellCabExtractPath -ItemType Directory -Force}
		
		# Pull down Dell XML CAB used in Dell Command Update, extract and Load
		Write-Verbose "Downloading Dell PC Index Cab"
		try {
			Invoke-WebRequestDownload -Uri "https://downloads.dell.com/catalog/CatalogIndexPC.cab" -OutFile $CabPathIndex -Verbose:$VerbosePreference
		} catch {
			Write-Error "Error downloading [https://downloads.dell.com/catalog/CatalogIndexPC.cab]"
			throw
		}
		If(Test-Path "$DellCabExtractPath\DellSDPCatalogPC.xml"){Remove-Item -Path "$DellCabExtractPath\DellSDPCatalogPC.xml" -Force}
		Start-Sleep -Seconds 1
		if (test-path $DellCabExtractPath){Remove-Item -Path $DellCabExtractPath -Force -Recurse}
		$null = New-Item -Path $DellCabExtractPath -ItemType Directory
		Write-Verbose "Expanding Dell PC Index Cab File..." 
		$null = expand $CabPathIndex $DellCabExtractPath\CatalogIndexPC.xml
		
		Write-Verbose "Loading Dell PC Index Catalog XML.... can take awhile"
		[xml]$XMLIndex = Get-Content "$DellCabExtractPath\CatalogIndexPC.xml"
		
		
		$SupportedModels = $XMLIndex.ManifestIndex.GroupManifest
		$SupportedModelsObject = @()
		foreach ($SupportedModel in $SupportedModels){
			$SPInventory = New-Object -TypeName PSObject
			$SPInventory | Add-Member -MemberType NoteProperty -Name "SystemID" -Value "$($SupportedModel.SupportedSystems.Brand.Model.systemID)" -Force
			$SPInventory | Add-Member -MemberType NoteProperty -Name "Model" -Value "$($SupportedModel.SupportedSystems.Brand.Model.Display.'#cdata-section')"  -Force
			$SPInventory | Add-Member -MemberType NoteProperty -Name "URL" -Value "$($SupportedModel.ManifestInformation.path)" -Force
			$SPInventory | Add-Member -MemberType NoteProperty -Name "Date" -Value "$($SupportedModel.ManifestInformation.version)" -Force		
			$SupportedModelsObject += $SPInventory 
		}
		
		if (($SupportedModelsObject | Measure).Count -le 0) {
			Write-Error "Error parsing Dell PC Index Catalog"
			throw
		}
	
		$script:_CachedSupportedModels = $SupportedModelsObject
	}
		
	return $SupportedModelsObject
}

function Find-DellBIOSUpdate {
	<#
		.SYNOPSIS
		Returns object with the latest available BIOS update.
		
		.DESCRIPTION
		Returns object with the latest available BIOS update. Can check given SKU or use the current system's SKU. Can return cached results in same session.
		
		.PARAMETER SKU
		Optional SKU to check. Will default to the current system's SKU if not given.
		
		.OUTPUTS
		Object with Path, BIOSVersion
		$null if no update available
		
		.NOTES
		You can lookup the SKU with (Get-CimInstance -ClassName Win32_ComputerSystem).SystemSKUNumber
		
		Path = URL path on http://downloads.dell.com/ to BIOS update.
	#>
	[CmdletBinding()]
	param(
		[string] $SKU
	)
	
	if (-Not [string]::IsNullOrEmpty($SKU)) {
		$SystemSKUNumber = $SKU
	} else {
		$SystemSKUNumber = (Get-CimInstance -ClassName Win32_ComputerSystem).SystemSKUNumber
	}
	if ([string]::IsNullOrEmpty($SystemSKUNumber)) {
		Write-Error "Invalid or empty system SKU number"
		throw
	}
	
	# Check cache first.
	$BIOSUpdates = $script:_CachedBIOSUpdates
	$NewestBIOSAvailable = $BIOSUpdates.$SystemSKUNumber
	if ($NewestBIOSAvailable -ne $null) {
		if ($NewestBIOSAvailable -eq $false) {
			Write-Verbose "Previous check attempt found in cache for SKU [$SystemSKUNumber], no BIOS available for this SKU"
			$NewestBIOSAvailable = $null
		} else {
			Write-Verbose "Latest BIOS available found in cache for SKU [$SystemSKUNumber]"
		}
	} else {
		$temproot = $script:_TEMPDIR
		if (!(Test-Path $script:_TEMPDIR)) { $null = New-Item -Path  $script:_TEMPDIR -ItemType Directory -Force }
		
		$CabPathIndexModel = "$temproot\DellCabDownloads\$SKU\CatalogIndexModel.cab"
		$DellCabExtractPath = "$temproot\DellCabDownloads\$SKU\DellCabExtract"
		if (!(Test-Path $DellCabExtractPath)){$null = New-Item -Path $DellCabExtractPath -ItemType Directory -Force}
		
		Write-Verbose "Using Dell Catalog to get Latest BIOS Version for SKU [$SystemSKUNumber]"
		try {
			$DellSKUObjs = Get-DellSupportedModels
		} catch {
			Write-Error $_
			throw
		}
		
		$DellSKUObj = $DellSKUObjs | Where-Object {$_.systemID -match $SystemSKUNumber} | Select-Object -First 1
		if (-Not $DellSKUObj) {
			Write-Error "Dell SKU [$SystemSKUNumber] not found in Dell PC Index Catalog"
			throw
		}
		$IndexURL = $DellSKUObj.URL
		if ([string]::IsNullOrWhitespace($IndexURL)) {
			Write-Error "Dell SKU [$SystemSKUNumber] has blank model index URL in Dell PC Index Catalog"
			throw
		}
		
		# Check cache first.
		$ModelIndexesXML = $script:_CachedModelIndexes
		if ($ModelIndexesXML.$IndexURL -ne $null) {
			$XMLIndexCAB = $ModelIndexesXML.$IndexURL
			if ($XMLIndexCAB -eq $false) {
				Write-Verbose "Previous invalid model index check attempt found in cache for SKU [$SystemSKUNumber], skipping"
				$XMLIndexCAB = $null
			}
		} else {
			Write-Verbose "Using Catalog URI [$IndexURL] from Model [$($DellSKUObj.Model)]"
			if (Test-Path $CabPathIndexModel){Remove-Item -Path $CabPathIndexModel -Force}
			try {
				Invoke-WebRequestDownload -Uri "http://downloads.dell.com/$IndexURL" -OutFile $CabPathIndexModel -Verbose:$VerbosePreference
			} catch {
				Write-Error "Error downloading [http://downloads.dell.com/$IndexURL]"
				throw
			}
			Write-Verbose "Extracting Dell Catalog for Model [$($DellSKUObj.Model)]"
			$null = expand $CabPathIndexModel $DellCabExtractPath\CatalogIndexPCModel.xml
			[xml]$XMLIndexCAB = Get-Content "$DellCabExtractPath\CatalogIndexPCModel.xml"
			
			# Save to cache
			if ($script:_CachedModelIndexes -isnot [hashtable]) {
				$script:_CachedModelIndexes = @{}
			}
			if ($XMLIndexCAB -ne $null) {
				$script:_CachedModelIndexes.Add($IndexURL, $XMLIndexCAB)
			} else {
				# Don't recheck for the same URL even if there are no results
				$script:_CachedModelIndexes.Add($IndexURL, $false)
			}
		}
		
		$NewestBIOSAvailable = $XMLIndexCAB.Manifest.SoftwareComponent | Where-Object {$_.ComponentType.value -eq "BIOS" -And -Not [string]::IsNullOrWhitespace($_.Path)} | Select Path,@{N="BIOSVersion"; E={ [Version]$_.vendorVersion }} | Sort -Descending BIOSVersion | Select -First 1
		
		# Save to cache
		if ($script:_CachedBIOSUpdates -isnot [hashtable]) {
			$script:_CachedBIOSUpdates = @{}
		}
		if ($NewestBIOSAvailable -ne $null) {
			$script:_CachedBIOSUpdates.Add($SystemSKUNumber, $NewestBIOSAvailable)
		} else {
			# Don't recheck for the same SKU even if there are no results
			$script:_CachedBIOSUpdates.Add($SystemSKUNumber, $false)
		}
	}
	
	if ($NewestBIOSAvailable) {
		Write-Verbose("Newest BIOS Version = [$($NewestBIOSAvailable.BIOSVersion)] for SKU [$SystemSKUNumber]")
		Write-Verbose("Download Path = [$($NewestBIOSAvailable.Path)]")
	}
				
	return $NewestBIOSAvailable
}

function Check-DellBIOSUpdate {
	<#
		.SYNOPSIS
		Returns object with Newest BIOS Available
		
		.DESCRIPTION
		Returns object with Newest BIOS Available. Can check given SKU or use the current system's SKU.
		
		.PARAMETER SKU
		Optional SKU to check. Will default to the current system's SKU if not given.
		
		.OUTPUTS
		Object with Path, BIOSVersion
		$null if no update available
		
		.NOTES
		You can lookup the SKU with (Get-CimInstance -ClassName Win32_ComputerSystem).SystemSKUNumber
		
		Path = URL path on http://downloads.dell.com/ to BIOS update.
	#>
	[CmdletBinding()]
	param(
		[string] $SKU
	)
	
	$UpdateBIOSAvailable = $null
	try {
		$Params = @{}
		if (-Not [string]::IsNullOrEmpty($SKU)) {
			$Params.Add('SKU', $SKU)
		}
		$NewestBIOSAvailable = Find-DellBIOSUpdate @Params -Verbose:$VerbosePreference
	} catch {
		Write-Error $_
		return -1
	}
	
	try {
		$CurrentBIOSVersion = [Version](Get-CIMInstance -Class Win32_BIOS -Property SMBIOSBIOSVersion -ErrorAction Stop | Select -ExpandProperty SMBIOSBIOSVersion)
	} catch {
		Write-Error "Error getting BIOS version"
		return -2
	}
	
	if ($CurrentBIOSVersion -lt $NewestBIOSAvailable.BIOSVersion) {
		$UpdateBIOSAvailable = $NewestBIOSAvailable
	} else {
		Write-Verbose "No BIOS Update available."
	}
	
	return $UpdateBIOSAvailable
}

function Download-DellBIOSUpdate {
	<#
		.SYNOPSIS
		Downloads latest available Dell BIOS.
		
		.DESCRIPTION
		Downloads latest available Dell BIOS. Can check given SKU or use the current system's SKU.
		
		.PARAMETER SKU
		Optional SKU to check. Will default to the current system's SKU if not given.
		
		.PARAMETER BIOSObject
		Optional BIOS object to download from Check-DellBIOSUpdate or Find-DellBIOSUpdate.
		
		.PARAMETER Path
		Optional directory path to download the BIOS file.
		
		.OUTPUTS
		Path to file if successful
		
		.NOTES
		Path = URL path on http://downloads.dell.com/ to BIOS update.
	#>
	[CmdletBinding(DefaultParameterSetName = 'None')]
	param( 
		[Parameter(Mandatory=$true, ParameterSetName = 'SKU')]
		[string] $SKU,
		
		[Parameter(Mandatory=$true, ParameterSetName = 'BIOSObject')]
		[PSObject] $BIOSObject,
		
		[Parameter(Mandatory=$false, ParameterSetName = 'SKU')]
		[Parameter(Mandatory=$false, ParameterSetName = 'BIOSObject')]
		[string] $Path
	)
	
	$NewestBIOSAvailable = $BIOSObject
	if (-Not $NewestBIOSAvailable) {
		Write-Verbose "Checking for Dell BIOS updates"
		try {
			if (-Not [string]::IsNullOrEmpty($SKU)) {
				$NewestBIOSAvailable = Find-DellBIOSUpdate -SKU $SKU -Verbose:$VerbosePreference
			} else {
				$NewestBIOSAvailable = Check-DellBIOSUpdate -Verbose:$VerbosePreference
			}
		} catch {
			Write-Error $_
			throw
		}
	}
	
	$DellBIOSDLFilePath = $null
	if ($NewestBIOSAvailable.BIOSVersion) {	
		if ([string]::IsNullOrEmpty($NewestBIOSAvailable.Path)) {
			Write-Error "Bad download path [$($NewestBIOSAvailable.Path)]"
			throw
		}
		$DownloadURI = "http://downloads.dell.com/$($NewestBIOSAvailable.Path)"
		if (-Not [string]::IsNullOrEmpty($Path)) {
			$DellBIOSDLPath = $Path
		} else {
			$temproot = $script:_TEMPDIR
			if (!(Test-Path $script:_TEMPDIR)) { $null = New-Item -Path  $script:_TEMPDIR -ItemType Directory -Force }
			
			$DellBIOSDLPath = "$temproot\DellBIOSDownloads"
		}
		
		if (!(Test-Path $DellBIOSDLPath)){$null = New-Item -Path $DellBIOSDLPath -ItemType Directory -Force}
		$BIOSFile = Split-Path -Leaf -Path $NewestBIOSAvailable.Path
		if ([string]::IsNullOrEmpty($BIOSFile)) {
			Write-Error "Bad download path [$($NewestBIOSAvailable.Path)]"
			throw
		}
		$DellBIOSDLFilePath = "$DellBIOSDLPath\$BIOSFile"
		if ((Test-Path $DellBIOSDLFilePath)) {
			Write-Host "Latest BIOS available: [$($NewestBIOSAvailable.BIOSVersion)]. [$DellBIOSDLFilePath] already exists, not re-downloading"
		} else {
			Write-Host "Latest BIOS available: [$($NewestBIOSAvailable.BIOSVersion)]. Downloading [$DownloadURI] to [$DellBIOSDLFilePath]"
			try {
				Invoke-WebRequestDownload -Uri $DownloadURI -OutFile $DellBIOSDLFilePath -Verbose:$VerbosePreference
			} catch {
				Write-Error "Error downloading [$DownloadURI]"
				throw
			}
		}
	}
	
	return $DellBIOSDLFilePath
}

function Install-DellBIOSUpdate {
	<#
		.SYNOPSIS
		Installs newest available Dell BIOS
		
		.DESCRIPTION
		Installs newest available Dell BIOS
		
		.PARAMETER BIOSPassword
		Password for BIOS update if needed
		
		.PARAMETER LogDir
		Optional directory for logging from BIOS Dell Update Package
				
		.OUTPUTS
		Exit code from BIOS Dell Update Package if successful
		-1 = Error checking for available BIOSes
		-2 = Error downloading BIOS file
		-3 = Error running update package
	#>
	[CmdletBinding()]
	param(
		[string] $BIOSPassword,
		
		[string] $LogDir
	)
	
	Write-Verbose "Checking for Dell BIOS updates"
	$NewestBIOSAvailable = $null
	try {
		$NewestBIOSAvailable = Check-DellBIOSUpdate -Verbose:$VerbosePreference
	} catch {
		Write-Error $_
		return -1
	}
	
	$exitCode = 0
	if ($NewestBIOSAvailable.BIOSVersion) {	
		try {
			$DellBIOSDLFilePath = Download-DellBIOSUpdate -BIOSObject $NewestBIOSAvailable -Verbose:$VerbosePreference
		} catch {
			Write-Error $_
			return -2
		}
		
		# Command-line arguments to the BIOS Dell Update Package
		# /s = silent
		# /bls = pause bitlocker until restart
		# /p=biospassword
		# /l="log filepath"
		$Arguments = "/s /bls"
		if (-Not [string]::IsNullOrEmpty($BIOSPassword)) {
			# Not sure if this can be quoted for spaces
			$Arguments += " /p=$BIOSPassword"
			Write-Verbose "BIOS password given"
		}
		if (-Not [string]::IsNullOrEmpty($LogDir)) {
			$logPath = "$($LogDir)\$($BIOSFile).log"
			$Arguments += " /l=`"$($LogDir)\$($BIOSFile).log`""
			if ((Test-Path $logPath -PathType Leaf)) { Clear-Content $logPath }
			Write-Verbose "Saving update package log to [$logPath]"
		}
		try {
			Write-Verbose("Calling : Start-Process -FilePath $DellBIOSDLFilePath -ArgumentList $Arguments -NoNewWindow -PassThru -ErrorAction Stop")
			$process = Start-Process -FilePath $DellBIOSDLFilePath -ArgumentList $Arguments -NoNewWindow -PassThru -ErrorAction Stop
			$handle = $process.Handle # Cache the process handle
			if ($process) {
				$null = $process | Wait-Process
				$exitCode = $process.ExitCode
			} else {
				Write-Error "Error calling Start-Process [$DellBIOSDLFilePath]"
				return -3
			}
		} catch {
			Write-Error $_
			return -3
		}
	} else {
		Write-Verbose "No updates available. Install skipped."
	}
	
	return $exitCode
}

# -- END FUNCTIONS --

# Main code
$_scriptName = split-path $PSCommandPath -Leaf	
if ([string]::IsNullOrEmpty($_scriptName)) {
	$_scriptName = "Get-DellBIOSUpdate"
}
$exitCode = -10

if (-Not [string]::IsNullOrEmpty($LogDir)) {
	$logFilename = "$($_scriptName).log"
	$logFilePath = "$LogDir\$logFileName"
	if (!(Test-Path $LogDir)){$null = New-Item -Path $LogDir -ItemType Directory -Force}
	Start-Transcript -Path $logFilePath -Force
}

# Set temporary directory used for extractions and downloads.
if (-Not [string]::IsNullOrEmpty($TempDir)) {
	$script:_TEMPDIR = $TempDir
} else {
	$script:_TEMPDIR = "${env:TEMP}\$_scriptName"
}

# Initialize cached objects
$script:_CachedSupportedModels = $null
$script:_CachedBIOSUpdates = @{}
$script:_CachedModelIndexes = @{}

$Params = @{}
If ($LogLevel -gt 0) {
	$Params.Add("Verbose", $true)
}

if ($Check -Or $PSCmdlet.ParameterSetName -eq 'Check') {
	try {
		if ($LogLevel -gt 0) {
			Write-Host "Calling Check-DellBIOSUpdate"
		}
		if (Check-DellBIOSUpdate @Params) {
			$exitCode = 1
		} else {
			$exitCode = 0
		}
	} catch {
		Write-Error $_
		$exitCode = -9
	}
} elseif ($Install) {
	$doInstall = $true
	#Make sure device is actually a Dell.
	$PCInfo = Get-CimInstance -Class Win32_ComputerSystem -Property Manufacturer
	if (-Not ($PCInfo.Manufacturer -like "Dell*" -OR $PCInfo.Manufacturer -like "Alienware*")) {
		Write-Error "Not a Dell system. Aborting install."
		$exitCode = -2
		$doInstall = $false
	}

	if (-Not $SkipPowerCheck) {
		Add-Type -Assembly System.Windows.Forms
		$PowerStatus = [System.Windows.Forms.SystemInformation]::PowerStatus
		if ($PowerStatus.PowerLineStatus -eq "Offline") {
			Write-Error "AC Power may be disconnected. Aborting install."
			$exitCode = -3
			$doInstall = $false
		}
	}

	if ($doInstall) {
		if (-Not [string]::IsNullOrEmpty($BIOSPassword)) {
			$Params.Add("BIOSPassword", $BIOSPassword)
		}
		if (-Not [string]::IsNullOrEmpty($LogDir)) {
			$Params.Add("LogDir", $LogDir)
		}
		try {
			if ($LogLevel -gt 0) {
				Write-Host "Calling Install-DellBIOSUpdate"
			}
			$exitCode = Install-DellBIOSUpdate @Params
			if ($LogLevel -gt 0) {
				Write-Host "Install-DellBIOSUpdate returned exit code [$exitCode]"
			}
			# 0 - Success, 2 - Restart required, 6 - Rebooting
			if ($exitCode -in 2,6) {
				$exitCode = 3010
			}
		} catch {
			Write-Error $_
			$exitCode = -9
		}
	}
} elseif ($Download) {
	#Make sure device is actually a Dell.
	$PCInfo = Get-CimInstance -Class Win32_ComputerSystem -Property Manufacturer
	if (-Not ($PCInfo.Manufacturer -like "Dell*" -OR $PCInfo.Manufacturer -like "Alienware*")) {
		Write-Error "Not a Dell system. Aborting download."
		$exitCode = -2
	} else {
		try {
			if (-Not [string]::IsNullOrWhitespace($Path)) {
				$Params.Add('Path', $Path)
			}
			if ($LogLevel -gt 0) {
				Write-Host "Calling Download-DellBIOSUpdate with params $($Params | ConvertTo-Json)"
			}
			$DellBIOSDLFilePath = Download-DellBIOSUpdate @Params
			if (-Not $LogLevel) {
				if ([string]::IsNullOrEmpty($DellBIOSDLFilePath)) {
					Write-Host "No BIOS Update available for this system"
				} else {
					Write-Host "BIOS Update downloaded to [$DellBIOSDLFilePath]"
				}
			}
		} catch {
			Write-Error $_
			$exitCode = -9
		}
	}
} elseif (-Not [string]::IsNullOrEmpty($DownloadSKU)) {
	try {
		$Params.Add('SKU', $DownloadSKU)
		if (-Not [string]::IsNullOrWhitespace($Path)) {
			$Params.Add('Path', $Path)
		}
		if ($LogLevel -gt 0) {
			Write-Host "Calling Download-DellBIOSUpdate with params $($Params | ConvertTo-Json)"
		}
		$DellBIOSDLFilePath = Download-DellBIOSUpdate @Params
		if (-Not $LogLevel) {
			if ([string]::IsNullOrEmpty($DellBIOSDLFilePath)) {
				Write-Host "No BIOS Update available for SKU [$DownloadSKU]"
			} else {
				Write-Host "BIOS Update for SKU [$DownloadSKU] available at [$DellBIOSDLFilePath]"
			}
		}
	} catch {
		Write-Error $_
		$exitCode = -9
	}
} elseif (-Not [string]::IsNullOrEmpty($DownloadCSV)) {
	$SKUs = $null
	try {
		$SKUs = Import-CSV $DownloadCSV
	} catch {
		Write-Error $_
		$exitCode = -4
	}
	$errorCount = 0
	foreach ($row in $SKUs) {
		try {
			$Params['SKU'] = $row.SKU
			if (-Not [string]::IsNullOrWhitespace($row.Path)) {
				$Params['Path'] = $row.Path
			}
			if ($LogLevel -gt 0) {
				Write-Host "Calling Download-DellBIOSUpdate with params $($Params | ConvertTo-Json)"
			}
			$DellBIOSDLFilePath = Download-DellBIOSUpdate @Params
			if (-Not $LogLevel) {
				if ([string]::IsNullOrEmpty($DellBIOSDLFilePath)) {
					Write-Host "No BIOS Update available for SKU [$($row.SKU)]"
				} else {
					Write-Host "BIOS Update for SKU [$($row.SKU)] available at [$DellBIOSDLFilePath]"
				}
			}
		} catch {
			Write-Error $_
			$errorCount++
		}
	}
	if ($errorCount -eq 0) {
		$exitCode = 0
	} else {
		$exitCode = -9
	}
} else {
	# Should never get here
	Write-Error "Invalid or unknown script parameters"
}

if ($LogLevel -gt 0) {
	Write-Host "Returning exit code: $exitCode"
}
Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

[System.Environment]::Exit($exitCode)