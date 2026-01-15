<#
	.SYNOPSIS
	Attempts to silently uninstall software given its name.
	
	.DESCRIPTION
	Attempts to silently uninstall software given its name from the registry. Usually this will be the same name as Add/Remove Programs. The script will throw an error if the uninstall keys still exists after an uninstallation attempt is done.
    
	.PARAMETER SoftwareName
	The display name of the software to uninstall (Required). This should be the same as the one from Add/Remove Programs. Can be a partial, in which case it will attempt to uninstall the first match found. 
	
	.PARAMETER Publisher
	Optional publisher to match along with display name. This can be partial.

	.PARAMETER Version
	Optional version to match. This is always an exact match.

	.PARAMETER ExactMatch
	Only accept exact matches for the given software and publisher.
	
	.PARAMETER Regex
	Treat SoftwareName and Publisher as regex patterns. Overriden by -ExactMatch switch.
	
	.PARAMETER Parameters
	Optional parameters added to what's given to the installer. Note if it's an EXE that's not MsiExec, it will assume an InstallShield with default additional parameters of "-uninst","-s". Use -OverrideParameters if you need to override these.
	
	.PARAMETER OverrideParameters
	If set, override the parameters given to setup instead of using defaults based on whats parsed from the registry. If it detects a Msiexec in the registry parameters then a parameter with "<MSI_GUID>" will contain the parsed GUID for the /X parameter.

	.PARAMETER ISSRecordResponseFile
	If set, record a response file to the given filename for InstallShield setup (required for full silent uninstall with some installers). You can then add it to your given parameters: -Parameters "-f1""Filename.is"""
	
	.PARAMETER Timeout
	Number of seconds to wait for the uninstall process to complete. Give null or 0 to wait indefinitely. Default: 3000
	
	.PARAMETER WaitInstallLocationDeleted
	If given, wait until the path specified in the InstallLocation registry entry is deleted by the called setup utility before exiting.
	
	.PARAMETER WaitInstallLocationDeletedTimeout
	Timeout in seconds to wait for the install location to be deleted. Default: 3000
	
	.PARAMETER OutputOnly
	If set, only output the found UninstallString before exiting. This is for debugging purposes.
	
	.EXAMPLE
	powershell.exe -NoProfile -Windowstyle Hidden -ExecutionPolicy Bypass -File "UninstallScriptPS.ps1" -SoftwareName "PowerFAIDS" -Parameters "-f1""%~dp0PowerFAIDs_Uninstall.iss""" -WaitInstallLocationDeleted

	.NOTES
	The script will attempt to parse and reuse the parameters from the registry. In the case of MsiExec, it will add default silent uninstall parameters. In the case of Exe files, it will give default InstallShield setup utility uninstall parameters. If the EXE is not either of those you will need to specify any needed silent uninstall parameters with the -OverrideParameters switch given. Ex: -Parameters "/s" -OverrideParameters
	
	For InstallShield setup files with prompts you can use the -ISSRecordResponseFile parameter to record a required response file for a fully automated uninstall.
	
	Make sure to test the uninstall outside of SCCM in case there's an uncaught prompt. If you still see a prompt, call the script using -OutputOnly to see what the parameters are. Then try opening up an admin command prompt and use common parameters like "setup.exe -?" or "setup.exe /?" to see if it displays a list of parameters (where setup.exe is the uninstaller filename). If that doesn't work, check online for any documentation on the uninstaller.

	Author: mcarras8
	
	Changelog
	01-09-2026 - mcarras8 - Strings used with -match are now properly escaped, added new parameter -Regex to override this behavior
	12-23-2025 - mcarras8 - Minor tweaks
	04-15-2025 - mcarras8 - Added support for QuietUninstallString. 
						  - Added support for passing exit codes.
						  - Script will now give a non-zero exit code if the install key still exists after the uninstall attempt.
#>

param(
	
	[parameter(Mandatory=$true, Position=0)]
	[string] $SoftwareName,
	
	[string] $Publisher,

	[string] $Version,

	[switch] $ExactMatch,
	
	[switch] $Regex,
	
	[string[]] $Parameters,
	
	[switch] $OverrideParameters,
	
	[string] $ISSRecordResponseFile,
		
	[int] $Timeout = 3000,
	
	[switch] $WaitInstallLocationDeleted,
	
	[int] $WaitInstallLocationDeletedTimeout = 3000,
	
	[switch] $OutputOnly
)

$_scriptName = split-path $PSCommandPath -Leaf

# Escape given strings unless we're using -ExactMatch or -Regex switches
$_displayname = $SoftwareName
$_publisher = $Publisher
if (-Not $ExactMatch -And -Not $Regex) {
	$_displayname = [Regex]::Escape($_displayname)
	$_publisher = [Regex]::Escape($_publisher)
}

# Check 32-bit, then 64-bit registry nodes.
$installLocation = $null
$uninstallString = $null
$hasQuietUninstallParams = $false
$installKey = gci "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { ( $ExactMatch -And $_.DisplayName -eq $_displayname -And ( [string]::IsNullOrEmpty($_publisher) -Or $_.Publisher -eq $_publisher ) ) -Or ( -Not $ExactMatch -And $_ -match $_displayname -And ( [string]::IsNullOrEmpty($_publisher) -Or $_.Publisher -match $_publisher ) ) -And ( [string]::IsNullOrEmpty($Version) -Or $_.DisplayVersion -eq $Version ) }
if ($installKey) {
	If (($installKey | Get-Member -Type NoteProperty | ? {$_.Name -eq 'InstallLocation'} | Measure).Count -gt 0) {
		$installLocation = $installKey | Select InstallLocation | Select -ExpandProperty InstallLocation
	}
	if (($installKey | Get-Member -Type NoteProperty | ? {$_.Name -eq 'QuietUninstallString'} | Measure).Count -gt 0 -And -Not [string]::IsNullOrWhitespace(($installKey | Select -ExpandProperty QuietUninstallString))) {
		$uninstallString = $installKey | Select QuietUninstallString | Select -ExpandProperty QuietUninstallString
		$hasQuietUninstallParams = $true
	} elseif (($installKey | Get-Member -Type NoteProperty | ? {$_.Name -eq 'UninstallString'} | Measure).Count -gt 0) {
		$uninstallString = $installKey | Select UninstallString | Select -ExpandProperty UninstallString
	}
}
# Check 64-bit registry
if ([string]::IsNullOrWhitespace($uninstallString)) {
	$installKey = gci "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" | foreach { gp $_.PSPath } | ? { ( $ExactMatch -And $_.DisplayName -eq $_displayname -And ( [string]::IsNullOrEmpty($_publisher) -Or $_.Publisher -eq $_publisher ) ) -Or ( -Not $ExactMatch -And $_ -match $_displayname -And ( [string]::IsNullOrEmpty($_publisher) -Or $_.Publisher -match $_publisher ) ) -And ( [string]::IsNullOrEmpty($Version) -Or $_.DisplayVersion -eq $Version ) }
	if ($installKey) {
		If (($installKey | Get-Member -Type NoteProperty | ? {$_.Name -eq 'InstallLocation'} | Measure).Count -gt 0) {
			$installLocation = $installKey | Select InstallLocation | Select -ExpandProperty InstallLocation
		}
		if (($installKey | Get-Member -Type NoteProperty | ? {$_.Name -eq 'QuietUninstallString'} | Measure).Count -gt 0 -And -Not [string]::IsNullOrWhitespace(($installKey | Select -ExpandProperty QuietUninstallString))) {
			$uninstallString = $installKey | Select QuietUninstallString | Select -ExpandProperty QuietUninstallString
			$hasQuietUninstallParams = $true
		} elseif (($installKey | Get-Member -Type NoteProperty | ? {$_.Name -eq 'UninstallString'} | Measure).Count -gt 0) {
			$uninstallString = $installKey | Select UninstallString | Select -ExpandProperty UninstallString
		}
	}
}

if ([string]::IsNullOrWhitespace($uninstallString)) {
	Write-Warning "[$_scriptName] Unable to find uninstall string for [$SoftwareName] in Registry, aborting"
	exit 0
}
Write-Host "[$_scriptname] DisplayName = [$($installKey.DisplayName)]"
Write-Host "[$_scriptname] Publisher = [$($installKey.Publisher)]"
Write-Host "[$_scriptname] DisplayVersion = [$($installKey.DisplayVersion)]"
Write-Host "[$_scriptname] UninstallString = [$uninstallString]"
Write-Host "[$_scriptname] InstallLocation = [$installLocation]"
If (-Not $OutputOnly) {
	# Get MSI path and parameters.
	$msiGUID = $null
	if ( $uninstallString -match '^MsiExec(?:\.exe) /[XIxi]\s?({[^}]+})') {
		if($ISSRecordResponseFile) {
			Write-Error "[$_scriptName] Parsed MsiExec setup is incompatible with -ISSRecordResponseFile"
			exit 1
		}
		$uninstallExe = 'Msiexec.exe'
		$msiGUID = $Matches[1]
		# Default parameters for Msiexec.
		$Params = @("/X", $msiGUID, "/qn", "/norestart", "REBOOT=REALLYSUPPRESS")
	} else {
		# Get EXE path and parameters.
		if ( $uninstallString -match '^"?([^"]+\.exe)"?(.+)') {
			$uninstallExe = $Matches[1]
			$Params = $Matches[2] -split " " | where {-not [string]::IsNullorWhitespace($_)}
			if($ISSRecordResponseFile) {
				$Params += @("-uninst","-r","-f1""$ISSRecordResponseFile""")
			} elseif (-Not $hasQuietUninstallParams) {
				$Params += @("-uninst","-s")
			}
		} else {
			# All other parameters.
			if ( $uninstallString -match '^"?([^"]+\.(?cmd|bat))"?(.+)') {
				if($ISSRecordResponseFile) {
					Write-Error "[$_scriptName] Parsed Cmd/Bat setup is incompatible with -ISSRecordResponseFile"
					exit 2
				}
				$uninstallExe = $Matches[1]
				$Params = $Matches[2] -split " " | where {-not [string]::IsNullorWhitespace($_)}
			}
		}
		# Additional parameters given to the setup utility.
		If ($Parameters) {
			If ($OverrideParameters) {
				If (-Not $msiGUID) {
					$Params = $Parameters
				} else {
					$Params = @()
					foreach($e in $Parameters) {
						if ($e -eq '<MSI_GUID>') {
							$Params += @($msiGUID)
						} else {
							$Params += @($e)
						}
					}
				}
			} else {
				$Params += $Parameters
			}
		}
	}

	if ([string]::IsNullOrWhitespace($uninstallExe)) {
		Write-Error "[$_scriptName] Unable to parse uninstall string for [$softwareName], aborting"
		exit 3
	}
	
	# & $uninstallExe $Params
	# Start uninstaller and wait for it to finish before continuing.
	$procParams = @{}
	If ($Params) {
		Write-Host("[$_scriptname] Calling: $uninstallExe {0}" -f ($Params -join " "))
		$procParams.Add("ArgumentList", $Params)
	} else {
		Write-Host("[$_scriptname] Calling: $uninstallExe")
	}
	$process = Start-Process -FilePath $uninstallExe @procParams -NoNewWindow -PassThru
	$handle = $process.Handle # Cache the process handle
	If ($Timeout) {
		$process | Wait-Process -Timeout $Timeout
	} else {
		$process | Wait-Process
	}
	$exitCode = $process.ExitCode
	
	# If option given, wait until the InstallLocation is gone before continuing.
	If($WaitInstallLocationDeleted -And -not [string]::IsNullOrWhitespace($installLocation)) {
		if ((Test-Path $installLocation -PathType Container)) {
			$msg = "[$_scriptName] InstallLocation [$installLocation] still found after uninstallation"
			if ($WaitInstallLocationDeletedTimeout -ne $null) {
				$msg += ", waiting up to $WaitInstallLocationDeletedTimeout seconds..."
			}
			Write-Host $msg
		}
		$sleepTime = 10
		$counter = 0
		if ($WaitInstallLocationDeletedTimeout -ne $null) {
			while($counter -le $WaitInstallLocationDeletedTimeout -And (Test-Path $installLocation -PathType Container)) {
				Start-Sleep $sleepTime
				$counter += $sleepTime
			}
		}
	}
	
	# Check to see if the registry key is gone, waiting up to 5 minutes.
	$maxSleepTime = 300
	$sleepTime = 10
	$counter = 0
	while ((Test-Path $installKey.PsPath) -And $counter -le $maxSleepTime) {
		Start-Sleep $sleepTime
		$counter += $sleepTime
	}
	if (Test-Path $installKey.PsPath) {
		Write-Error "Install key still exists after uninstallation attempt: $($installKey.PSPath)"
		exit 4
	}
	
	# Return process exit code
	exit $exitCode
}