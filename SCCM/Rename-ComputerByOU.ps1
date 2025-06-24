<#
    .SYNOPSIS
    Renames a computer to match its current sub-OU.
    
	.DESCRIPTION
    Renames a computer to match its current sub-OU. Assumes the renamed computer will be in the form NameStart-DeptCode-Serial.
    
	.PARAMETER RootDN
    The distinguished name for the root OU to search for sub-OUs.
	
	.PARAMETER DeptCodeRegex
	The regular expression to extract the dept code from the current sub-OU. Defaults to '([A-Za-z0-9]{2,3})', which allows either a 2 character or 3 character dept code. Assumes OUs are in the form of "Division-DeptCode" and existing computers are in the form of either "Division-DeptCode-Serial" or "DeptCode-Serial".
	
	.PARAMETER NameSeparator
	The character used to separate different parts of the name. Defaults to a dash ('-').
	
	.PARAMETER RenameStart
	Overrides the start of the new name before the first dash.
	
	.PARAMETER SystemSerial
	Grab the serial from the system, rather than from the last of the name. If the system serial is too long it will default back to the last part of the name.
	
	.PARAMETER OnlyEthernet
	Only update name when connected to Ethernet, failing otherwise. Can be overriden by -VPNIPRegex.
	
	.PARAMETER VPNIPRegex
	Allow updating over wifi if any of the adapter IP addresses matches the given regex. Default: 10.247.
	
	.PARAMETER Restart
	Have the Rename-Computer cmdlet restart the system.
	
	.PARAMETER NoExit
	Do not exit at the end.
	
	.PARAMETER LogFilePath
	The full filepath for the log file. Defaults to C:\Windows\CCM\Logs\RenameComputerByOU.log if not given.
	
    .NOTES
	
	Error codes:
	0 - Success
	1 - Not connected to Ethernet and -OnlyEthernet is given
	2 - System.DirectoryServices.DirectorySearcher failure finding current computer
	3 - Cannot find dept code in current OU
	4 - Name parsing error or bad serial
	5 - Rename computer failure
	
    Author: MJC 3-11-24
#>
[CmdletBinding(DefaultParameterSetName = 'RootDN')]
param(
	[Parameter(Mandatory=$true, ParameterSetName='RootDN', Position=0)]
	[alias('RootDistinguishedName')]
	[ValidateNotNullOrEmpty()]
	[string]$RootDN, 
	
	[ValidateNotNullOrEmpty()]
	[string]$DeptCodeRegex='([A-Za-z0-9]{2,3})',
	
	[ValidatePattern('^[^,]+')]
	[string]$NameSeparator='-',
	
	[string]$RenameStart,
	
	[switch]$SystemSerial,
	
	[switch]$OnlyEthernet,
	
	[AllowEmptyString()]
	[string]$VPNIPRegex="10\.247\.",
	
	[switch]$Restart,
	
	[switch]$NoExit,
	
	[string]$LogFilePath
)

# Defaults.
$_scriptName = "Rename-ComputerByOU.ps1"
$maxComputerNameLength = 15

# Should remain 0 unless errors are caught.
$exitCode = 0

# Set current computer name
$_computerName = ${ENV:COMPUTERNAME}

# Directory for logs.
# Saves to C:\Windows\CCM\Logs by default.
if([string]::IsNullOrWhitespace($LogFilePath)) {
	try {
		$logDir = (Get-ItemProperty -Path HKLM:\Software\Microsoft\CCM\Logging\@Global).LogDirectory
	} catch {
		Write-Error $_
	}
	
	If([string]::IsNullOrWhitespace($logDir)) {
		$logDir = "${ENV:SystemDrive}\Temp"
		
		# If local path for logs doesn't exist, create it
		If (!(Test-Path $logDir)) { 
			New-Item -Path $logDir -Type Directory -Force 
		}
	}
	$logFileName = "RenameComputerByOU.log"
	$_logFilePath = "$logDir\$logFileName"
}

Clear-Content -Path $_logFilePath

# Check if using Ethernet if the -OnlyEthernet switch is given.
# Also check if we're on the VPN network, which will override -OnlyEthernet.
If ($OnlyEthernet -And (Get-NetAdapter -Physical | where {$_.Status -eq 'Up' -And ($_.PhysicalMediaType -eq "802.3" -Or $_.InterfaceDescription -match "Ethernet")} | Measure-Object | Select -ExpandProperty Count) -le 0) {
	$onVPN = $false
	if (-Not [string]::IsNullOrEmpty($VPNIPRegex)) {
		$allOnlineAdapters = Get-NetAdapter | where {$_.Status -eq "Up"}
		$onVPN = (Get-NetIPAddress -AddressFamily IPv4 | where {$_.ifIndex -in $allOnlineAdapters.ifIndex -And $_.IPAddress -match $VPNIPRegex} | Measure-Object).Count -gt 0
	}
	if ($onVPN) {
		Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] System is on WiFi with valid VPN IP, continuing" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName) -PassThru
	} else {
		Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] ERROR: Not connected to Ethernet or VPN and -OnlyEthernet switch was given." -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName) -PassThru
		$exitCode = 1
	}
}

If($exitCode -eq 0) {
	# Get domain controller DN from Root
	$DC = $RootDN.Substring($RootDN.IndexOf('DC='))

	# Get computer's current OU using DirectorySearcher
	try {
		$searcher = New-Object System.DirectoryServices.DirectorySearcher
		$searcher.SearchScope = 'Subtree'
		$searcher.SearchRoot = [ADSI]"LDAP://$DC"
		$searcher.filter = "(&(objectCategory=computer)(objectClass=computer)(cn=$_computerName))"
		$computerDN = $searcher.FindOne().Properties.distinguishedname
		$computerOU = $computerDN.Substring($($computerDN).IndexOf('OU='))
	} catch {
		Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] ERROR: {2}" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName, $_.Exception.Message)
		Write-Error $_
		$exitCode = 1
	}
	Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] computerDN=$computerDN" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName)
	Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] computerOU=$computerOU" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName)

	# Get new dept code.
	$newDeptCode = $null
	if($computerOU -match "^OU=[^$NameSeparator]+[$NameSeparator]$DeptCodeRegex[^,]*,") {
		$newDeptCode = $Matches[1]
	}
	if( [string]::IsNullOrWhitespace($newDeptCode) ) {
		Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] ERROR: Could not find valid deptCode from DN [{2}]" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName, $computerDN) -PassThru
		$exitCode = 2
	}
	Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] newDeptCode=$newDeptCode" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName)

	# Get current dept code and serial from name.
	$nameStart = $null
	$currDeptCode = $null
	$currSerial = $null
	if(($namePieces = $_computerName -split $NameSeparator) -And $namePieces.Count -gt 0) {
		If($namePieces.Count -eq 2) {
			$currDeptCode = $namePieces[0]
			$currSerial = $namePieces[1]
		} elseif ($namePieces.Count -eq 3) {
			$nameStart = $namePieces[0]
			$currDeptCode = $namePieces[1]
			$currSerial = $namePieces[2]
		} else {
			Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] Could not parse name -- returned {2} pieces using separator '{3}'" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName, $namePieces.Count, $NameSeparator) -PassThru
			$exitCode = 4
		}
	}
	Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] nameStart=$nameStart, currDeptCode=$currDeptCode, currSerial=$currSerial" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName)

	# Get current serial from WMI if missing.
	$serial = $null
	If ($SystemSerial) {
		$serial = (Get-WmiObject -class win32_bios).SerialNumber
	} else {
		$serial = $currSerial
	}
	Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] serial=$serial" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName)

	If(-Not [string]::IsNullOrWhitespace($RenameStart)) {
		$newComputerName = $RenameStart
	} else {
		$newComputerName = $nameStart
	}
	# Double-check name start is valid.
	If([string]::IsNullOrWhitespace($newComputerName)) {
		Write-Error "Name start is missing or could not be parsed, aborting"
		$exitCode = 4
	} else {
		# If using $serial is too long, default back to whatever currSerial is
		$tentativeComputerName = "$newComputerName$NameSeparator$newDeptCode$NameSeparator$serial"
		If($tentativeComputerName.Length -gt $maxComputerNameLength) {
			If([string]::IsNullOrWhitespace($currSerial)) {
				Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] ERROR: Computer name [{2}] would be too long with system serial [{3}] and could not parse serial in current name, aborting" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName, $newComputerName, $serial)
				$exitCode = 4
			} else {
				$tentativeComputerName = "$newComputerName$NameSeparator$newDeptCode$NameSeparator$currSerial"
			}
		}
		$newComputerName = $tentativeComputerName
	}
	if ($exitCode -eq 0) {
		If ([string]::IsNullOrWhitespace($newComputerName) -Or $newComputerName.Length -gt $maxComputerNameLength ) {
			Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] ERROR: Invalid or missing new computer name [[{2}], aborting" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName, $newComputerName) -PassThru
			$exitCode = 4
		} else {
			If ($newComputerName -eq $_computerName) {
				Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] Current computer name already matches new computed name, no change needed" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName)
			} else {
				try {
					Rename-Computer -NewName $newComputerName -Force -Restart:$Restart
					Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] Computer will be renamed [$newComputerName] on next restart" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName)
				} catch {
					Add-Content -Path $_logFilePath -Value ("[{0}] [{1}] ERROR: {2}" -f (Get-Date).toString("yyyy/MM/dd HH:mm:ss"), $_scriptName, $_.Exception.Message)
					Write-Error $_
					$exitCode = 5
				}
			}
		}
	}
}

if (-Not $NoExit) {
	# Return error code (0 by default)
	exit $exitCode
}