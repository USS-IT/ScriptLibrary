<#
	.SYNOPSIS
	Checks or installs latest Dell BIOS for this system.
	
	.DESCRIPTION
	Checks or installs latest Dell BIOS for this system.
	
	.PARAMETER Install
	Installs the latest BIOS if it's an update.
	
	.PARAMETER SkipPowerCheck
	Skips the AC Power check when installing BIOSes.
	
	.PARAMETER BIOSPassword
	Password to BIOS to perform update, if set
	
	.PARAMETER CheckTSVar
	Call Check-BIOSUpdate and save the result as true/false in the given Task Sequence Variable.
	
	.PARAMETER TempDir
	Directory to use for temporary files. Will create directory if not found. Default is $env:TEMP\Get-DellBIOSUpdate
	
	.PARAMETER LogDir
	Directory to store / look for log files. Will create directory if not found. Defaults to C:\Dell\Logs
	
	.PARAMETER LogLevel
	Level of console logging.
	
	.OUTPUTS
	Exit code 0 - Without -Install, No BIOS update available
	Exit code 1 - Without -Install, BIOS update available
	Exit code 3010 - With -Install, BIOS update successful, restart required
	Exit code -1 - Function thrown error
	Exit code -2 - Uncaught or unexpected error
	All other exit codes - Returned from BIOS Dell Update Package (see notes)
	
	.EXAMPLE
	Get-DellBIOSUpdate.ps1 -Install
	
	.NOTES
	Author: Matt Carras (mcarras8) 6-17-26
	
	Some code copied from https://github.com/gwblok/garytown/blob/master/hardware/Dell/CommandUpdate/EMPS/Dell-EMPS.ps1
	
	You can check for exit codes using _SMSTSLastActionRetCode in a task sequence.
	
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

param( 
	[switch] $Install,
	
	[switch] $SkipPowerCheck,
	
	[string] $BIOSPassword,
	
	[string] $CheckTSVar,
		
	[string] $TempDir,
	
	[AllowEmptyString()]
	[string] $LogDir = "C:\Dell\Logs",
	
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
		Returns array of supported models from Dell's CatalogIndexPC.cab
	
		.OUTPUTS
		Supported models with SystemID, Model, URL, and Date
		
		.NOTES
		SystemID = System SKU.
		URL = Path on http://downloads.dell.com/ to CatalogIndexPCModel.xml. This contains all the driver and software downloads for that SKU-specific Model.
	#>
	[CmdletBinding()]
	param()
	
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
	return $SupportedModelsObject
}

function Check-DellBIOSUpdate {
	<#
		.SYNOPSIS
		Returns object with Newest BIOS Available
		
		.DESCRIPTION
		Returns object with Newest BIOS Available
		
		.OUTPUTS
		Object with Path, BIOSVersion
		$null if no update available
		
		.NOTES
		Path = URL path on http://downloads.dell.com/ to BIOS update.
	#>
	[CmdletBinding()]
	param()
	
	$temproot = $script:_TEMPDIR
	if (!(Test-Path $script:_TEMPDIR)) { $null = New-Item -Path  $script:_TEMPDIR -ItemType Directory -Force }
	
	$CabPathIndexModel = "$temproot\DellCabDownloads\CatalogIndexModel.cab"
	$DellCabExtractPath = "$temproot\DellCabDownloads\DellCabExtract"
	if (!(Test-Path $DellCabExtractPath)){$null = New-Item -Path $DellCabExtractPath -ItemType Directory -Force}
			
	$SystemSKUNumber = (Get-CimInstance -ClassName Win32_ComputerSystem).SystemSKUNumber
	Write-Verbose "Using Dell Catalog to get Latest BIOS Version for SKU [$SystemSKUNumber]"
	$DellSKUs = Get-DellSupportedModels
	if (($DellSKUs | Measure).Count -le 0) {
		Write-Error "Error parsing Dell PC Index Catalog"
		throw
	}
	$DellSKU = $DellSKUs | Where-Object {$_.systemID -match $SystemSKUNumber} | Select-Object -First 1
	if (-Not $DellSKU) {
		Write-Error "Dell SKU [$SystemSKUNumber] not found in Dell PC Index Catalog"
		throw
	}
	Write-Verbose "Using Catalog URI [$($DellSKU.URL)] from Model [$($DellSKU.Model)]"
	if (Test-Path $CabPathIndexModel){Remove-Item -Path $CabPathIndexModel -Force}
	try {
		Invoke-WebRequestDownload -Uri "http://downloads.dell.com/$($DellSKU.URL)" -OutFile $CabPathIndexModel -Verbose:$VerbosePreference
	} catch {
		Write-Error "Error downloading [http://downloads.dell.com/$($DellSKU.URL)]"
		throw
	}
	Write-Verbose "Extracting Dell Catalog for Model [$($DellSKU.Model)]"
	$null = expand $CabPathIndexModel $DellCabExtractPath\CatalogIndexPCModel.xml
	[xml]$XMLIndexCAB = Get-Content "$DellCabExtractPath\CatalogIndexPCModel.xml"
	
	$NewestBIOSAvailable = $XMLIndexCAB.Manifest.SoftwareComponent | Where-Object {$_.ComponentType.value -eq "BIOS" -And -Not [string]::IsNullOrWhitespace($_.Path)} | Select Path,@{N="BIOSVersion"; E={ [Version]$_.vendorVersion }} | Sort -Descending BIOSVersion | Select -First 1
	
	$CurrentBIOSVersion = [Version](Get-CIMInstance -Class Win32_BIOS -Property SMBIOSBIOSVersion | Select -ExpandProperty SMBIOSBIOSVersion)
	
	Write-Host "Current BIOS version: $CurrentBIOSVersion"
	Write-Host "Newest BIOS version: $($NewestBIOSAvailable.BIOSVersion)"
	if ($CurrentBIOSVersion -lt $NewestBIOSAvailable.BIOSVersion) {
		return $NewestBIOSAvailable
	}
	Write-Host "No BIOS Update available."
	
	return $null
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
		-2 = Bad download path from available BIOSes
		-3 = Error downloading BIOS file
		-4 = Error 
		
		.NOTES
		Path = URL path on http://downloads.dell.com/ to BIOS update.
	#>
	[CmdletBinding()]
	param(
		[string] $BIOSPassword,
		
		[string] $LogDir
	)
	
	$exitCode = 0
	Write-Verbose "Checking for Dell BIOS updates"
	$NewestBIOSAvailable = $null
	try {
		$NewestBIOSAvailable = Check-DellBIOSUpdate
	} catch {
		Write-Error $_
		return -1
	}
	if ($NewestBIOSAvailable.BIOSVersion) {	
		Write-Verbose("BIOSVersion = $($NewestBIOSAvailable.BIOSVersion)")
		Write-Verbose("Path = $($NewestBIOSAvailable.Path)")
		if ([string]::IsNullOrEmpty($NewestBIOSAvailable.Path)) {
			Write-Error "Bad download path [$($NewestBIOSAvailable.Path)]"
			return -2
		}
		$temproot = $script:_TEMPDIR
		if (!(Test-Path $script:_TEMPDIR)) { $null = New-Item -Path  $script:_TEMPDIR -ItemType Directory -Force }
		$DownloadURI = "http://downloads.dell.com/$($NewestBIOSAvailable.Path)"
		$DellBIOSDLPath = "$temproot\DellBIOSDownloads"
		if (!(Test-Path $DellBIOSDLPath)){$null = New-Item -Path $DellBIOSDLPath -ItemType Directory -Force}
		$BIOSFile = Split-Path -Leaf -Path $NewestBIOSAvailable.Path
		if ([string]::IsNullOrEmpty($BIOSFile)) {
			Write-Error "Bad download path [$($NewestBIOSAvailable.Path)]"
			return -2
		}
		$DellBIOSDLFilePath = "$DellBIOSDLPath\$BIOSFile"
		Write-Host "Updated BIOS available: [$($NewestBIOSAvailable.BIOSVersion)]. Downloading [$DownloadURI] to [$DellBIOSDLFilePath]"
		try {
			Invoke-WebRequestDownload -Uri $DownloadURI -OutFile $DellBIOSDLFilePath -Verbose:$VerbosePreference
		} catch {
			Write-Error "Error downloading [$DownloadURI]"
			return -3
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
		}
		if (-Not [string]::IsNullOrEmpty($LogDir)) {
			$logPath = "$($LogDir)\$($BIOSFile).log"
			$Arguments += " /l=`"$($LogDir)\$($BIOSFile).log`""
			if ((Test-Path $logPath -PathType Leaf)) { Clear-Content $logPath }
		}
		
		Write-Verbose("Calling : Start-Process -FilePath $DellBIOSDLFilePath -ArgumentList $Arguments -NoNewWindow -PassThru")
		$process = Start-Process -FilePath $DellBIOSDLFilePath -ArgumentList $Arguments -NoNewWindow -PassThru
		$handle = $process.Handle # Cache the process handle
		if ($process) {
			$null = $process | Wait-Process
			$exitCode = $process.ExitCode
		} else {
			Write-Error "Error calling Start-Process [$DellBIOSDLFilePath]"
			return -4
		}
	} else {
		Write-Verbose "No updates available. Install skipped."
	}
	
	return $exitCode
}

# -- END FUNCTIONS --

# Main code
$_scriptName = split-path $PSCommandPath -Leaf
$exitCode = -2

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
	$script:_TEMPDIR = "${env:TEMP}\Get-DellBIOSUpdate"
}
			
#Make sure device is actually a Dell.
$PCInfo = Get-CimInstance -Class Win32_ComputerSystem -Property Manufacturer
if (-Not ($PCInfo.Manufacturer -like "Dell*" -OR $PCInfo.Manufacturer -like "Alienware*")) {
	Write-Error "Not a Dell system. Aborting."
	[System.Environment]::Exit(-1)
}

if (-Not $SkipPowerCheck) {
	Add-Type -Assembly System.Windows.Forms
	$PowerStatus = [System.Windows.Forms.SystemInformation]::PowerStatus
	if ($PowerStatus.PowerLineStatus -eq "Offline") {
		Write-Error "AC Power may be disconnected. Aborting."
		[System.Environment]::Exit(-2)
	}
}

$Params = @{}
If ($LogLevel -gt 0) {
	$Params.Add("Verbose", $true)
}
if ($Install) {
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
		if ($exitCode -ne $null -And $exitCode -in 0,2,6) {
			$exitCode = 3010
		}
	} catch {
		Write-Error $_
		if ($LogLevel -gt 0) {
			Write-Host "Install-DellBIOSUpdate errored with exit code [$exitCode]"
		}
		$exitCode = -1
	}
} elseif (-Not [string]::IsNullOrEmpty($CheckTSVar)) {
	try {
		$tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
		if ($LogLevel -gt 0) {
			Write-Host "Calling Check-DellBIOSUpdate"
		}
		if (Check-DellBIOSUpdate @Params) {
			$tsenv.Value($CheckTSVar) = "true"
		} else {
			$tsenv.Value($CheckTSVar) = "false"
		}
		if ($LogLevel -gt 0) {
			Write-Host "Saved value in TSVar [$CheckTSVar]"
		}
		$exitCode = 0
	} catch {
		$exitCode = -1
	}
} else {
	if ($LogLevel -gt 0) {
		Write-Host "Calling Check-DellBIOSUpdate"
	}
	if (Check-DellBIOSUpdate @Params) {
		$exitCode = 1
	} else {
		$exitCode = 0
	}
}

Stop-Transcript -ErrorAction SilentlyContinue | Out-Null

[System.Environment]::Exit($exitCode)