<#
    .SYNOPSIS
    Moves computer object into established sub-OU.
    
	.DESCRIPTION
    Moves a computer object into a sub-OU that matches the beginning the computer name. Example: USS-IT-ABC123 would be moved into "USS-IT" sub-OU, if it exists at $RootDN.
    
	.PARAMETER RootDN
    The distinguished name for the root OU to search for sub-OUs. Either this or DestDN are required.
	
	.PARAMETER DestDN
    The distinguished name for destination OU. Either this or RootDN are required.
	
	.PARAMETER ComputerName
    Optional computer name. ${ENV:COMPUTERNAME} is used if not given if working outside of WinPE task sequence.
	
	.PARAMETER InTaskSequence
	Assumes we're running in a task sequence (and have access to the TSEnvironment COM object). Note this may cause problems if you run the task as a user other than SYSTEM.
	
	.PARAMETER DoNotExit
	Do not exit at the end.
	
    .NOTES
	Give -Verbose for more details.
	
	This script can work within a SCCM OSD task sequence. While the script can be run within WinPE with -InTaskSequence switch, it's recommended to wait until after it reboots into the newly applied OS.
	
	The script must be run as an account with the correct permissions to query and move computer objects in the given OUs.
	
	Error codes:
	0 - Success
	1 - System.DirectoryServices.DirectorySearcher failure finding current computer
	2 - System.DirectoryServices.DirectorySearcher failure searching for sub-OUs
	3 - ADSI moveto failure
	
    Author: MJC 10-12-2022
#>
[CmdletBinding(DefaultParameterSetName = 'RootDN')]
param(
	[Parameter(Mandatory=$true, ParameterSetName='RootDN', Position=0)]
	[alias('RootDistinguishedName')]
	[string]$RootDN, 
	
	[Parameter(Mandatory=$true, ParameterSetName='DestDN', Position=0)]
	[alias('DestinationDistinguishedName')]
	[string]$DestDN, 
	
	[string]$ComputerName,
	
	[alias('InTaskSequence')]
	[switch]$InTS,
	
	[switch]$DoNotExit
)

# Should remain 0 unless errors are caught.
$exitCode = 0

# Check to see if we're in WinPE.
$InWinPE = (Test-Path -Path Registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlset\Control\MiniNT)

# Try to grab values from the task sequence environment
$tsenv_OSDComputerName = $null
$tsenv__SMSTSMachineName = $null
if ($InTS) {
	$tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
	$InWinPE = $InWinPE -Or ($tsenv.Value("_SMSTSInWinPE") -eq $true)
	$tsenv_OSDComputerName = $tsenv.Value("OSDComputerName")
	$tsenv__SMSTSMachineName = $tsenv.Value("_SMSTSMachineName")
}

# Set current computer name, if not given previously
if (-Not [string]::IsNullOrWhitespace($ComputerName)) {
	$_computerName = $ComputerName
} elseif (-Not $InWinPE) {
	$_computerName = ${ENV:COMPUTERNAME}
} elseif (-Not [string]::IsNullOrWhitespace($tsenv_OSDComputerName)) {
	$_computerName = $tsenv_OSDComputerName
} else {
	$_computerName = $tsenv__SMSTSMachineName
}

# Get domain controller DN from either Destination or Root
if (-Not [string]::IsNullOrEmpty($DestDN)) {
	$DC = $DestDN.Substring($DestDN.IndexOf('DC='))
} else {
	$DC = $RootDN.Substring($RootDN.IndexOf('DC='))
}
	
# Get computer's current OU using DirectorySearcher
try {
	$searcher = New-Object System.DirectoryServices.DirectorySearcher
	$searcher.SearchScope = 'Subtree'
	$searcher.SearchRoot = [ADSI]"LDAP://$DC"
	$searcher.filter = "(&(objectCategory=computer)(objectClass=computer)(cn=$_computerName))"
	$computerDN = $searcher.FindOne().Properties.distinguishedname
	$computerOU = $computerDN.Substring($($computerDN).IndexOf('OU='))
} catch {
	Write-Error $_
	$exitCode = 1
}

$newOU = $null
if ($exitCode -eq 0 -And -Not [string]::IsNullOrEmpty($computerDN) -And -Not [string]::IsNullOrEmpty($computerOU)) {	
	if (-Not [string]::IsNullOrEmpty($DestDN)) {
		$newOU = $DestDN
	} else {
		# Get all sub-OUs starting at $RootDN using DirectorySearcher (if $DestDN is not given)
		try {
			$searcher.SearchScope = 'Subtree'
			$searcher.SearchRoot = [ADSI]"LDAP://$RootDN"
			$searcher.filter = '(objectClass=organizationalUnit)'
			foreach($ou in $searcher.FindAll()) {
				$dn = $ou.Properties.distinguishedname
				$name = $ou.Properties.name
				if (-Not [string]::IsNullOrEmpty($dn) -And -Not [string]::IsNullOrEmpty($name) -And $dn -ne $RootDN) {
					if ($_computerName -like "$name-*") {
						$newOU = $dn
						break
					}
				}
			}
		} catch {
			Write-Error $_
			$exitCode = 2
		}
	}
	
	# Use ADSI to move computer object, if necessary
	if ($exitCode -eq 0 -And -Not [string]::IsNullOrEmpty($newOU)) {
		if ($newOU -eq $computerOU) {
			Write-Verbose("[Move-ADComputerToSubOU.ps1] AD computer is already in destination OU [$newOU], not moving")
		} else {
			try {
				Write-Verbose("[Move-ADComputerToSubOU.ps1] Moving [$computerDN] to destination OU [$newOU]")
				# Bind to computer object
				$computer = [ADSI]"LDAP://$computerDN"
				# Bind to target OU and then move computer object
				$computer.psbase.MoveTo([ADSI]"LDAP://$newOU")
			} catch {
				Write-Error $_
				$exitCode = 3
			}
		}
	}
}

if (-Not $DoNotExit) {
	# Return error code (0 by default)
	[System.Environment]::Exit($exitCode)
}