#Script: Invoke-WUAPS.ps1
#Created: 2026-02-16
#Author: mcarras8@jh.edu
#Description: Search for and install updates from Windows Update without a restart. To be used during OSD.
#Notes:

#Requires -RunAsAdministrator
$ErrorActionPreference = 'Stop'

# -- For Task Sequence Environment --
$tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
$tsProgressUI = New-Object -ComObject Microsoft.SMS.TSProgressUI

# --- Create a WUA session ---
$session = New-Object -ComObject Microsoft.Update.Session
$session.ClientApplicationID = "Invoke-WUAPS"

$searcher = $session.CreateUpdateSearcher()

# Criteria examples:
#   Software updates only: "IsInstalled=0 and Type='Software'"
#   Add IsHidden=0 to skip hidden updates
$criteria = "IsInstalled=0 and IsHidden=0 and Type='Software'"

try {
	$tsProgressUI.ShowActionProgress(
		$tsenv.Value("_SMSTSOrgName"),
		$tsenv.Value("_SMSTSPackageName"),
		$tsenv.Value("_SMSTSCustomProgressDialogMessage"),
		$tsenv.Value("_SMSTSCurrentActionName"),
		[Convert]::ToUInt32($tsenv.Value("_SMSTSNextInstructionPointer")),
		[Convert]::ToUInt32($tsenv.Value("_SMSTSInstructionTableSize")),
		"Searching for updates...",
		0,
		100
	)
} catch {}

Write-Host "Searching for updates with criteria: $criteria"
$searchResult = $searcher.Search($criteria)

if ($searchResult.Updates.Count -eq 0) {
    Write-Host "No applicable updates found."
    exit 0
}

# --- Build collection of updates to download/install ---
$toInstall = New-Object -ComObject Microsoft.Update.UpdateColl

for ($i=0; $i -lt $searchResult.Updates.Count; $i++) {
    $u = $searchResult.Updates.Item($i)

    # Auto-accept EULA if needed
    if (-not $u.EulaAccepted) { $u.EulaAccepted = $true }

    # Optional: skip "Preview" updates (common in OSD)
    if ($u.Title -match 'Preview') { 
        Write-Host "Skipping Preview update: $($u.Title)"
        continue
    }

    [void]$toInstall.Add($u)
    Write-Host "Selected: $($u.Title)"
}

if ($toInstall.Count -eq 0) {
    Write-Host "Nothing selected after filtering."
    exit 0
}
	
# --- Download ---
$downloader = $session.CreateUpdateDownloader()
$downloader.Updates = $toInstall

try {
	$tsProgressUI.ShowActionProgress(
		$tsenv.Value("_SMSTSOrgName"),
		$tsenv.Value("_SMSTSPackageName"),
		$tsenv.Value("_SMSTSCustomProgressDialogMessage"),
		$tsenv.Value("_SMSTSCurrentActionName"),
		[Convert]::ToUInt32($tsenv.Value("_SMSTSNextInstructionPointer")),
		[Convert]::ToUInt32($tsenv.Value("_SMSTSInstructionTableSize")),
		"Downloading $($toInstall.Count) updates...",
		50,
		100
	)
} catch {}

Write-Host "Downloading $($toInstall.Count) update(s)..."
$downloadResult = $downloader.Download()
Write-Host "Download ResultCode: $($downloadResult.ResultCode)"

# --- Install (quiet, no forced reboot) ---
$installer = $session.CreateUpdateInstaller()
$installer.Updates = $toInstall
$installer.ForceQuiet = $true

# Optional: this property exists on the installer; read-only per Microsoft
# It indicates if reboot is required *before* installing updates
# https://learn.microsoft.com/windows/win32/api/wuapi/nf-wuapi-iupdateinstaller-get_rebootrequiredbeforeinstallation
$preReboot = $installer.RebootRequiredBeforeInstallation
Write-Host "RebootRequiredBeforeInstallation: $preReboot"  # informational only [2](https://learn.microsoft.com/en-us/windows/win32/api/wuapi/nf-wuapi-iupdateinstaller-get_rebootrequiredbeforeinstallation)

try {
	$tsProgressUI.ShowActionProgress(
		$tsenv.Value("_SMSTSOrgName"),
		$tsenv.Value("_SMSTSPackageName"),
		$tsenv.Value("_SMSTSCustomProgressDialogMessage"),
		$tsenv.Value("_SMSTSCurrentActionName"),
		[Convert]::ToUInt32($tsenv.Value("_SMSTSNextInstructionPointer")),
		[Convert]::ToUInt32($tsenv.Value("_SMSTSInstructionTableSize")),
		"Installing $($toInstall.Count) updates...",
		75,
		100
	)
} catch {}

Write-Host "Installing updates..."
$installResult = $installer.Install()

Write-Host "Install ResultCode: $($installResult.ResultCode)"
Write-Host "Install RebootRequired: $($installResult.RebootRequired)"

try {
	$tsProgressUI.ShowActionProgress(
		$tsenv.Value("_SMSTSOrgName"),
		$tsenv.Value("_SMSTSPackageName"),
		$tsenv.Value("_SMSTSCustomProgressDialogMessage"),
		$tsenv.Value("_SMSTSCurrentActionName"),
		[Convert]::ToUInt32($tsenv.Value("_SMSTSNextInstructionPointer")),
		[Convert]::ToUInt32($tsenv.Value("_SMSTSInstructionTableSize")),
		"Cleaning up...",
		100,
		100
	)
} catch {}

# --- Also check global reboot-required status via WUA SystemInfo ---
$sysInfo = New-Object -ComObject Microsoft.Update.SystemInfo
$rebootRequired = $sysInfo.RebootRequired

Write-Host "SystemInfo.RebootRequired: $rebootRequired"

# Exit codes you can key off in a Task Sequence:
# 0 = no updates or installed without reboot needed
# 3010 = success, reboot required (common convention)
$exitCode = 0
if ($rebootRequired -or $installResult.RebootRequired) {
    $exitCode = 3010
}
[System.Environment]::Exit($exitCode)