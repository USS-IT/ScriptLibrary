<#
    .SYNOPSIS
    Renames the current computer.
    
	.DESCRIPTION
    Renames the current computer, checking for a valid name and forcing a restart.
    
	.PARAMETER NewName
    Required. The new name for the computer.
	
	.PARAMETER ComputerName
    Optional computer name. ${ENV:COMPUTERNAME} is used if not given if working outside of WinPE task sequence.
	
	.PARAMETER InTaskSequence
	Assumes we're running in a task sequence (and have access to the TSEnvironment COM object)
	
	.PARAMETER Restart
	Restarts the computer after renaming.
	
	.PARAMETER DoNotExit
	Do not exit at the end.
	
    .NOTES
	Give -Verbose for more details.
	
	This script can work within a SCCM OSD task sequence, but must be run outside of WinPE (after restarting in newly applied OS).
	
	Error codes:
	0 - Success
	1 - Invalid name
	2 - Error returned from Rename-Computer
	
    Author: MJC 10-12-2022
#>
param(
	[Parameter(Mandatory=$true)]
	[string]$NewName, 
	
	[string]$ComputerName,
	
	[switch]$Restart,
	
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
If (-Not [string]::IsNullOrWhitespace($ComputerName)) {
	$curCompName = $ComputerName
} ElseIf (-Not $InWinPE) {
	$curCompName = ${ENV:COMPUTERNAME}
} ElseIf (-Not [string]::IsNullOrWhitespace($tsenv_OSDComputerName)) {
	$curCompName = $tsenv_OSDComputerName
} Else {
	$curCompName = $tsenv__SMSTSMachineName
}

# Check variable is not blank
If (-Not [string]::IsNullOrWhitespace($NewName)) {
	# Check to see if name is same as current
	If ($NewName -eq $curCompName) {
		Write-Verbose "[Rename-ComputerEx.ps1] Computer already named [$NewName], skipping rename"
	} ElseIf ($InWinPE) {
		Write-Warning "[Rename-ComputerEx.ps1] This script cannot rename a computer in WinPE"
	} Else {
		# Invalid NetBIOS names:
		# * Over 15 characters
		# * Certain 1 and 2 character names (limited here to 3+)
		# * Cannot include symbols (except - and _, though underscore may cause DNS issues)
		# * Cannot start with a special character or end with a minus
		# * Cannot only contain digits (limited here to require at least 1 letter)
		If ($NewName -notmatch "^(?=.*[a-zA-Z].*)[a-zA-Z0-9][a-zA-Z0-9_-]{1,13}[a-zA-Z0-9_]$") {
			Write-Error "[Rename-ComputerEx.ps1] New computer name is invalid [$NewName]"
			$exitCode = 1
		} Else {
			try {
				Rename-Computer -NewName $NewName -Force -Restart:$Restart
				Write-Verbose "[Rename-ComputerEx.ps1] Computer will be renamed [$NewName] on next restart"
			} catch {
				Write-Error $_
				$exitCode = 2
			}
		}
	}
}

if (-Not $DoNotExit) {
	# Return error code (0 by default)
	[System.Environment]::Exit($exitCode)
}