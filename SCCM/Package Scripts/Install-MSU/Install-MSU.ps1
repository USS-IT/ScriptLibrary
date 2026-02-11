<#
.SYNOPSIS
    Install the given Windows Update MSU package and check for successful installation afterwards
.PARAMETER MSU
	Full path to the Windows Update Package to install
.PARAMETER CheckExitCode
    Optional. Check for a valid exit code from wusa.exe.
.PARAMETER InstalledPackage
	Optional. Full name of installed package according to Get-WindowsPackage
.PARAMETER TimeoutMinutes
    Optional. Maximum time in minutes to keep checking installation. Default is 15.
.PARAMETER LogPath
    Optional. Path to create script log file. Default is C:\TEMP\Install-MSU.log.
.NOTES
	Author: Matt Carras (mcarras8)
	Created: 2-6-26
#>
param (
    [Parameter(Mandatory=$true)]
    [string]$MSU,
	
	[Parameter(Mandatory=$false)]
    [switch]$CheckExitCode,
	
	[Parameter(Mandatory=$false)]
    [string]$InstalledPackage,
	
	[Parameter(Mandatory=$false)]
	[int]$TimeoutMinutes=15,
	
	[Parameter(Mandatory=$false)]
	[string]$LogPath="C:\TEMP\Install-MSU.log"
)

# wusa.exe exit codes - successly installed
$installSuccessCodes = @(
	0,           # Installed
    2359301,     # Installed, reboot required
    2359302      # Already installed
)
# wusa.exe exit codes - not applicable
$notApplicableCodes = @(
	2359303,     # Already uninstalled
	2149842967,	 # Update not applicable
	2149842966,  # Installation not allowed by policy
	-2145124329  # Update not applicable
)
$allSuccessCodes = $installSuccessCodes + $notApplicableCodes

Start-Transcript $LogPath -Force

$fpMSU = $MSU
if ($fpMSU -notlike "*\*") {
	$fpMSU = $PSScriptRoot + "\" + $MSU
}
If (-Not(Test-Path $fpMSU -PathType Leaf)) {
	Write-Error "Package to install [$fpMSU] does not exist"
	exit -1
}
try {
	Write-Host "** Waiting for TrustedInstaller and TiWorker to finish if they're already running"
	Get-Process -Name TrustedInstaller -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -Name TiWorker -ErrorAction SilentlyContinue | Wait-Process
	
	# Start installation. This will likely spawn child processes (TrustedInstaller, TiWorker, etc.).
	Write-Host("** $env:SystemRoot\System32\wusa.exe " + '"' + $fpMSU + '"' + "/quiet /norestart")
	$process = Start-Process -FilePath "$env:SystemRoot\System32\wusa.exe" -ArgumentList @(('"' + $fpMSU + '"'),"/quiet","/norestart") -NoNewWindow -PassThru -Wait
	# Saving the handle will allow us to query the exit code later.
	$handle = $process.Handle
	Write-Host "** wusa.exe ExitCode: $($process.ExitCode)"
	
	Write-Host "** Waiting for TrustedInstaller and TiWorker to finish after installation attempt"
	Get-Process -Name TrustedInstaller -ErrorAction SilentlyContinue | Wait-Process
	Get-Process -Name TiWorker -ErrorAction SilentlyContinue | Wait-Process
} catch {
	throw $_
}

# Check for successful installation
if ($CheckExitCode -And $process.ExitCode -notin $allSuccessCodes) {
	Write-Error "wusa.exe returned invalid exitcode [$($process.ExitCode)]"
	exit -2
} elseif (-Not [string]::IsNullOrEmpty($InstalledPackage)) {
	if ($process.ExitCode -in $notApplicableCodes) {
		Write-Host "** Update not applicable, skipping installation detection"
	} else {
		Write-Host "** Checking for [$InstalledPackage] installation every minute for up to $TimeoutMinutes minutes"
		$counter = 0
		$isInstalled = $false
		while (-not $isInstalled -And $counter -lt $TimeoutMinutes) {
			try {
				$pkg = Get-WindowsPackage -Online -PackageName $InstalledPackage
				$isInstalled = ($pkg.PackageState -ne $null)
			} catch {}
			$counter++
			Start-Sleep -Seconds 60
		}
		if (-not $isInstalled) {
			Write-Error "$InstalledPackage not found after MSU installation"
			exit -3
		}
	}
}

Stop-Transcript

exit 0