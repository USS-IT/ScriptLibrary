<#
    .SYNOPSIS
    Sets the default computer name in an imaging task sequence.
    
	.DESCRIPTION
    Sets the default computer name in an imaging task sequence. This is mainly needed if run from TSMedia.
	
	.PARAMETER DefaultNamePrefix
    The default prefix to go before the serial #. Only used if we can't read the name from the registry.
	
	.PARAMETER MaxSerialLength
	Maximum length of serial # before truncation. Only used if we can't read the name from the registry.
	
    .NOTES
	mcarras8 7-24-25
#>
param(
	[string]$DefaultNamePrefix="USS-XX-", 
	
	[int]$MaxSerialLength=7
)

$tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
$OSDComputerName = $tsenv.Value("OSDComputerName")
$_SMSTSMachineName = $tsenv.Value("_SMSTSMachineName")
$serial = $tsenv.Value("_SMSTSSerialNumber")

# Only continue if we have an invalid OSDComputerName.
If($OSDComputerName -like "MININT-*") {
	# First, see if we can load the old name from the system's registry.
	$regcompname = $null
	$regfile = "${ENV:OSDWindowsDriveLetter}\Windows\system32\config\system"
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

	# If we have a valid name from the system's old registry, use that. Otherwise, generate one from the serial.
	if (-Not [string]::IsNullOrWhitespace($regcompname) -And $regcompname -notlike "MININT-*") {
		$default_computer_name = $regcompname
	} else {
		# Build a default name if no valid name form registry
		# Get the serial from WMI if not set in TS variable.
		if ([string]::IsNullOrWhitespace($serial)) {
			$serial = Get-WMIObject -Class Win32_BIOS | Select -ExpandProperty SerialNumber
			if (-Not [string]::IsNullOrWhitespace($serial)) {
				throw "Invalid serial number"
			}
		}

		if ($serial.Length -gt $MaxSerialLength) {
			$serial = $serial.Substring(0,$MaxSerialLength)
		}

		$default_computer_name = $DefaultNamePrefix + $serial
	}

	# Set the OSDComputerName TS variable for joining
	$tsenv.Value("OSDComputerName") = $default_computer_name
}

[System.Environment]::Exit(0)