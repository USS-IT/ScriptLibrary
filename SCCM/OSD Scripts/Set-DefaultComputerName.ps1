<#
    .SYNOPSIS
    Sets the default computer name in an imaging task sequence.
    
	.DESCRIPTION
    Sets the default computer name in an imaging task sequence. This is mainly needed if run from TSMedia / USB Standalone.
	
	.PARAMETER InvalidNameRegex
	Will only look at computers MATCHing regex pattern. By default, looks for names starting with "MIN" or missing a dash: "^(MIN.+|[^-]+)$". Give an empty string to skip this check. 
	
	.PARAMETER DefaultNamePrefix
    The default prefix to go before the serial #. Only used if we can't read the name from the registry. Note NetBIOS has limit of 15 characters total.
	
	.PARAMETER MaxSerialLength
	Maximum serial length allowed before its truncated. Default: 7
	
	.PARAMETER CheckRegistry
	Attempt to check the previously imaged system's registry for the current name.
	
    .NOTES
	Author: Matt Carras (mcarras8)
	Created: 7-24-25
#>
param(
	[string]$InvalidNameRegex="^(MIN.+|[^-]+)$",

	[string]$DefaultNamePrefix="USS-XX-", 
	
	[int]$MaxSerialLength=7,
	
	[switch]$CheckRegistry
)

$tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
$OSDComputerName = $tsenv.Value("OSDComputerName")
$_SMSTSMachineName = $tsenv.Value("_SMSTSMachineName")
$serial = $tsenv.Value("_SMSTSSerialNumber")
$OSDDriveLetter = $tsenv.Value("OSDWindowsDriveLetter")

# Only continue if we have an invalid OSDComputerName.
# Found some cases where OSDComputerName may be set to "0".
If([string]::IsNullOrWhitespace($OSDComputerName) -Or $OSDComputerName -eq "0" -Or (-Not [string]::IsNullOrWhitespace($InvalidNameRegex) -And $OSDComputerName -match $InvalidNameRegex)) {
	$regcompname = $null
	
	# First, see if we can load the old name from the system's registry.
	if ($CheckRegistry) {
		if ([string]::IsNullOrEmpty($OSDDriveLetter)) {
			$OSDDriveLetter = "C:"
		}
		$regfile = "$OSDDriveLetter\Windows\system32\config\system"
		if ((Test-Path $regfile -PathType Leaf)) {
			& reg.exe load HKLM\TempHive $regfile
		} else {
			$regfile = "C:\Windows\system32\config\system"
			if ((Test-Path $regfile -PathType Leaf)) {
				& reg.exe load HKLM\TempHive $regfile
			} else {
				$regfile = $null
			}
		}
		if ($regfile -ne $null) {
			$regcompname = (Get-ItemProperty -Path "HKLM:\TempHive\ControlSet001\Control\ComputerName\ComputerName" -Name Computername).Computername
			& reg.exe unload HKLM\TempHive
		}
	}

	# If we have a valid name from the system's old registry, use that. Otherwise, generate one from the serial.
	if (-Not [string]::IsNullOrWhitespace($regcompname) -And $regcompname -ne "0" -and $regcompname -notmatch $InvalidNameRegex) {
		$default_computer_name = $regcompname
		Write-Host("[Set-DefaultComputerName.ps1] Found [$default_computer_name] from registry")
	} else {
		# Build a default name if no valid name form registry
		# Get the serial from WMI if not set in TS variable.
		if ([string]::IsNullOrWhitespace($serial)) {
			$serial = Get-WMIObject -Class Win32_BIOS | Select -ExpandProperty SerialNumber
			if (-Not [string]::IsNullOrWhitespace($serial)) {
				throw "Invalid serial number"
			}
		}

		# Check if serial needs to be truncated
		if ($serial.Length -gt $MaxSerialLength) {
			$serial = $serial.Substring(0,$MaxSerialLength)
		}

		# NetBIOS has a max of 15 characters
		if ($default_computer_name.Length -gt 15) {
			Write-Host("[Set-DefaultComputerName.ps1] Truncating [$default_computer_name]")
			$default_computer_name = $default_computer_name.Substring(0,15)
		}
	}

	if (-Not [string]::IsNullOrWhitespace($default_computer_name) -And $default_computer_name -ne "0") {
		# Set the OSDComputerName TS variable for joining
		$tsenv.Value("OSDComputerName") = $default_computer_name
		Write-Host("[Set-DefaultComputerName.ps1] Setting OSDComputerName to [$default_computer_name]")
	}
} else {
	Write-Host("[Set-DefaultComputerName.ps1] OSDComputerName is valid [$OSDComputerName]")
}

[System.Environment]::Exit(0)