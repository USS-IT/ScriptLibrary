<#
.SYNOPSIS
    Collects all relevant Windows event log entries from the upgrade window and exports them to .evtx and .csv files.

.DESCRIPTION
    Queries all event logs relevant to Windows Setup, driver activity, PnP, kernel boot,
    Windows Update, and general system/application errors for the given window.

    Outputs:
      - One merged .csv with all events (human-readable, importable to Excel)
      - One .evtx per queried log channel (openable in Event Viewer)
      - A summary .txt with counts and any flagged high-interest events

.PARAMETER OutputDir
    Directory to write all output files to.
	
.PARAMETER MaxEventsPerLog
    Safety cap on events per log channel to prevent runaway output.
    Default: 2000

.PARAMETER StartDT
    Start date & time to parse logs.
	Default: Last 4 hours

.PARAMETER NoCSV
	Don't output events to a merged CSV file.

.PARAMETER NoEVTX
	Don't export .evtx event files.

.PARAMETER NoSummary
	Don't output a summary of findings.
	
.NOTES
    Tested on 5.1 and 7.2.
    Requires wevtutil.exe (built into Windows) for .evtx export.
	
	Author: Matt Carras (mcarras8)
	Created: 8-12-2026
#>

param(
	[Parameter(Mandatory=$true)]
	[string] $OutputDir,
	
    [int]    $MaxEventsPerLog = 2000,
	[string] $StartDT,
	
	[switch] $NoCSV,
	[switch] $NoEVTX,
	[switch] $NoSummary
)

#Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# ---------------------------------------------------------------------------
# 0. Logging helper
# ---------------------------------------------------------------------------
function Write-Log {
    param(
		[Parameter(Mandatory=$true, Position=0)]
		[string]$Message,
		
		[Parameter(Mandatory=$false, Position=1)]
		[string]$Level = "INFO"
	)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$ts][$Level] $Message"
    Write-Output $entry
    if ($script:LogFileHandle) {
        Add-Content -Path $script:ScriptLog -Value $entry -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
# 1. Prepare output directory
# ---------------------------------------------------------------------------
if (!(Test-Path $OutputDir)) {
	try {
		$null = New-Item -ItemType Directory -Path $OutputDir -Force
	} catch {
		Write-Error "FATAL: Could not create output directory '$OutputDir': $_"
		exit 1
	}
}
	
$timestamp       = Get-Date -Format "yyyyMMdd_HHmmss"
$hostname        = $env:COMPUTERNAME
$script:ScriptLog = "$OutputDir\Save-UpgradeEventLogs_$timestamp.log"
$script:LogFileHandle = $true
Add-Content -Path $script:ScriptLog -Value "" -ErrorAction SilentlyContinue  # touch the file

Write-Log "=== Save-UpgradeEventLogs.ps1 started ==="
Write-Log "Hostname    : $hostname"
Write-Log "OutputDir   : $OutputDir"

# ---------------------------------------------------------------------------
# 2. Resolve the upgrade start timestamp
# ---------------------------------------------------------------------------
$startTimeStr = $StartDT

# Parse the timestamp
$startTime = $null
if ($startTimeStr) {
    try {
        $startTime = [datetime]::Parse($startTimeStr, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
        Write-Log "Parsed start time (UTC): $($startTime.ToString('o'))"
    } catch {
        Write-Log "Could not parse timestamp '$startTimeStr': $_" "WARN"
    }
}

# Ultimate fallback: last 4 hours (catches most upgrade scenarios)
if (-not $startTime) {
    $startTime = (Get-Date).ToUniversalTime().AddHours(-4)
    Write-Log "WARNING: No start timestamp found. Falling back to last 4 hours: $($startTime.ToString('o'))" "WARN"
}

$endTime = (Get-Date).ToUniversalTime()
Write-Log "Collection window: $($startTime.ToString('o'))  -->  $($endTime.ToString('o'))"

# ---------------------------------------------------------------------------
# 3. Define event log channels to query
# ---------------------------------------------------------------------------
# Each entry: LogName, optional EventIds filter, Description
$logSources = @(

    # Core setup and upgrade channels
    @{ Log = "Setup";                                               Ids = @();          Desc = "Windows Setup" }
    @{ Log = "System";                                              Ids = @();          Desc = "System (all levels)" }
    @{ Log = "Application";                                         Ids = @(1,2,3);     Desc = "Application (Crit/Err/Warn)" }

    # Windows Update / WUA
    @{ Log = "Microsoft-Windows-WindowsUpdateClient/Operational";   Ids = @();          Desc = "Windows Update Client" }

    # PnP / Driver activity — most relevant for 0xC1900101
    @{ Log = "Microsoft-Windows-Kernel-PnP/Configuration";          Ids = @();          Desc = "Kernel PnP Configuration" }
    @{ Log = "Microsoft-Windows-UserPnp/DeviceInstall";             Ids = @();          Desc = "User PnP Device Install" }
    @{ Log = "Microsoft-Windows-DriverFrameworks-UserMode/Operational"; Ids = @();      Desc = "Driver Frameworks (UserMode)" }
    @{ Log = "Microsoft-Windows-SetupAPI/Dev";                      Ids = @();          Desc = "SetupAPI Dev" }
    @{ Log = "Microsoft-Windows-SetupAPI/Operational";              Ids = @();          Desc = "SetupAPI Operational" }

    # Kernel boot / BCD / bootmgr
    @{ Log = "Microsoft-Windows-Kernel-Boot/Operational";           Ids = @();          Desc = "Kernel Boot" }
    @{ Log = "Microsoft-Windows-Kernel-General/Operational";        Ids = @();          Desc = "Kernel General" }

    # Setup platform (Windows 10/11 feature update orchestration)
    @{ Log = "Microsoft-Windows-SetupPlatform/Operational";         Ids = @();          Desc = "SetupPlatform" }
    @{ Log = "Microsoft-Windows-SetupPlatform/Admin";               Ids = @();          Desc = "SetupPlatform Admin" }

    # CBS / Component Store (TrustedInstaller — staging issues)
    @{ Log = "Microsoft-Windows-Servicing/Operational";             Ids = @();          Desc = "Windows Servicing (CBS)" }

    # Secure Boot / UEFI CA
    @{ Log = "Microsoft-Windows-Kernel-EventTracing/Admin";         Ids = @();          Desc = "Kernel Event Tracing Admin" }
    @{ Log = "Microsoft-Windows-SecureBoot/Operational";            Ids = @();          Desc = "Secure Boot" }

    # Deployment / DISM
    @{ Log = "Microsoft-Windows-Deployment-Services-Diagnostics/Debug"; Ids = @();     Desc = "Deployment Services" }
    @{ Log = "Microsoft-Windows-DISM/Operational";                  Ids = @();          Desc = "DISM Operational" }

    # Reliability / WER
    @{ Log = "Application";                                         Ids = @(1001,1002); Desc = "WER/Crash reports (App)" }

    # BitLocker (relevant for Secure Boot CA interaction)
    @{ Log = "Microsoft-Windows-BitLocker/BitLocker Management";    Ids = @();          Desc = "BitLocker Management" }
    @{ Log = "Microsoft-Windows-BitLocker-DrivePreparationTool/Operational"; Ids = @(); Desc = "BitLocker Drive Prep" }
)

# Remove duplicates (Application appears twice intentionally with different Id filters;
# merge them into separate queries - deduplicate by Log+Ids combo)
$seenKeys = @{}
$logSources = $logSources | Where-Object {
    $key = "$($_.Log)|$($_.Ids -join ',')"
    if ($seenKeys.ContainsKey($key)) { return $false }
    $seenKeys[$key] = $true
    return $true
}

# ---------------------------------------------------------------------------
# 4. Query events and build master collection
# ---------------------------------------------------------------------------
$allEvents   = [System.Collections.Generic.List[PSObject]]::new()
$channelStats = [System.Collections.Generic.List[PSObject]]::new()

if (-Not $NoCSV) {
	Write-Log "--- Beginning event log collection across $(($logSources | Measure).Count) channels ---"

	foreach ($src in $logSources) {

		$logName = $src.Log
		$desc    = $src.Desc
		$ids     = $src.Ids

		# Build filter hashtable
		$filter = @{
			LogName   = $logName
			StartTime = $startTime.ToLocalTime()   # Get-WinEvent uses local time for StartTime
			EndTime   = $endTime.ToLocalTime()
		}
		if ($ids -and ($ids | Measure).Count -gt 0) {
			$filter["Id"] = $ids
		}

		$count = 0
		try {
			$events = Get-WinEvent -FilterHashtable $filter -MaxEvents $MaxEventsPerLog -ErrorAction Stop | Sort TimeCreated

			$count = ($events | Measure).Count

			foreach ($ev in $events) {
				$msgRaw = ""
				try { $msgRaw = $ev.Message } catch { $msgRaw = "[Message unavailable]" }
				$msgShort = ($msgRaw -replace "`r`n|`r|`n", " ").Trim()
				if ($msgShort.Length -gt 1000) { $msgShort = $msgShort.Substring(0, 1000) + "..." }

				$allEvents.Add([PSCustomObject]@{
					TimeCreatedUTC = $ev.TimeCreated.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
					LogName        = $ev.LogName
					Level          = $ev.LevelDisplayName
					EventId        = $ev.Id
					ProviderName   = $ev.ProviderName
					TaskDisplayName= $ev.TaskDisplayName
					Message        = $msgShort
					RecordId       = $ev.RecordId
					Channel        = $desc
				})
			}

			Write-Log "  [$desc] [$logName]: $count events collected."
		} catch [System.Exception] {
			$errMsg = $_.Exception.Message
			if ($errMsg -like "*No events were found*" -or $errMsg -like "*There are no more files*") {
				Write-Log "  [$desc] [$logName]: 0 events in window."
				$count = 0
			} elseif ($errMsg -like "*does not exist*" -or $errMsg -like "*not found*" -or $errMsg -like "*cannot be opened*") {
				Write-Log "  [$desc] [$logName]: Channel not available on this system." "WARN"
				$count = -1
			} else {
				Write-Log "  [$desc] [$logName]: Query error - $errMsg" "WARN"
				$count = -1
			}
		}

		$channelStats.Add([PSCustomObject]@{
			Channel   = $desc
			LogName   = $logName
			EventCount = $count
		})
	}

	Write-Log "Total events collected across all channels: $(($allEvents | Measure).Count)"

	# ---------------------------------------------------------------------------
	# 5. Export merged CSV (sorted by time)
	# ---------------------------------------------------------------------------
	$csvPath = "$OutputDir\${hostname}_UpgradeEvents_${timestamp}.csv"
	try {
		$allEvents | Sort-Object TimeCreatedUTC | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
		Write-Log "Merged CSV exported: $csvPath [$(($allEvents | Measure).Count) rows]"
	} catch {
		Write-Log "ERROR: Could not write CSV: $_" "ERROR"
	}
}

# ---------------------------------------------------------------------------
# 6. Export individual .evtx files via wevtutil
# ---------------------------------------------------------------------------
if (-Not $NoEVTX) {
	Write-Log "--- Exporting .evtx files via wevtutil ---"

	# Format timestamps for wevtutil XPath query (local time, XML datetime format)
	$xpathStart = $startTime.ToLocalTime().ToString("yyyy-MM-ddTHH:mm:ss")
	$xpathEnd   = $endTime.ToLocalTime().ToString("yyyy-MM-ddTHH:mm:ss")

	# Unique log names only (we query the channel once for .evtx regardless of Id filter)
	$uniqueLogs = $logSources | % {$_.Log} | Select -Unique

	foreach ($logName in $uniqueLogs) {

		# Sanitize log name for use as a filename
		$safeLogName = $logName -replace "[/\\:*?`"<>|]", "_"
		$evtxPath    = "$OutputDir\${hostname}_${safeLogName}_${timestamp}.evtx"

		# Build XPath 1.0 query to filter by time window
		$xpath = "*[System[TimeCreated[@SystemTime >= '$xpathStart' and @SystemTime `<= '$xpathEnd']]]"

		try {
			$wevtArgs = @(
				"epl",          # export-log
				$logName,
				$evtxPath,
				"/q:$xpath",
				"/ow:true"      # overwrite if exists
			)
			$result = & wevtutil.exe $wevtArgs 2>&1
			if ($LASTEXITCODE -eq 0) {
				$size = (Get-Item $evtxPath -ErrorAction SilentlyContinue).Length
				Write-Log "  Exported: $safeLogName `($size bytes`)"
			} else {
				Write-Log "  wevtutil failed for '$logName' `(exit $LASTEXITCODE`): $result" "WARN"
				# Clean up empty/failed file
				Remove-Item $evtxPath -Force -ErrorAction SilentlyContinue
			}
		} catch {
			Write-Log "  wevtutil exception for '$logName': $_" "WARN"
		}
	}
}

# ---------------------------------------------------------------------------
# 7. Flag high-interest events for the summary
# ---------------------------------------------------------------------------

if (-not $NoCSV) {
	# Patterns that are especially relevant to 0xC1900101-0x20017
	$highInterestPatterns = @(
		"0xC1900101",
		"0x20017",
		"0x80070020",           # ERROR_SHARING_VIOLATION
		"0x80070005",           # ERROR_ACCESS_DENIED
		"0x80070002",           # ERROR_FILE_NOT_FOUND
		"0x800B0100",           # TRUST_E_NOSIGNATURE
		"sharing violation",
		"access denied",
		"boot critical",
		"BootCritical",
		"SPSVCINST_BOOTMGR",
		"SafeOS",
		"SAFE_OS",
		"Cleanup external drivers",
		"Rollback",
		"rollback",
		"SetupPlatform",
		"setup failed",
		"cannot be found",
		"timed out",
		"timeout",
		"UEFICA2023",
		"Secure Boot",
		"BitLocker recovery",
		"crash",
		"dump"
	)

	$flaggedEvents = $allEvents | Where-Object {
		$msg = $_.Message
		$flagged = $false
		foreach ($p in $highInterestPatterns) {
			if ($msg -like "*$p*") { $flagged = $true; break }
		}
		$flagged
	}

	Write-Log "High-interest events flagged: $(($flaggedEvents | Measure).Count)"
}

# ---------------------------------------------------------------------------
# 8. Write summary report
# ---------------------------------------------------------------------------
if (-Not $NoSummary) {
	$summaryPath = "$OutputDir\${hostname}_UpgradeSummary_${timestamp}.txt"
	$summaryLines = [System.Collections.Generic.List[string]]::new()

	$summaryLines.Add("=" * 70)
	$summaryLines.Add("  Windows 25H2 Upgrade Event Log Collection Summary")
	$summaryLines.Add("=" * 70)
	$summaryLines.Add("")
	$summaryLines.Add("Hostname      : $hostname")
	$summaryLines.Add("Collection run: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') UTC")
	$summaryLines.Add("Window start  : $($startTime.ToString('o'))")
	$summaryLines.Add("Window end    : $($endTime.ToString('o'))")
	$summaryLines.Add("Total events  : $(($allEvents | Measure).Count)")
	$summaryLines.Add("Flagged events: $(($flaggedEvents | Measure).Count)")
	$summaryLines.Add("")

	# OS / hardware info
	$os = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue
	$cs = Get-CimInstance Win32_ComputerSystem  -ErrorAction SilentlyContinue
	$bios = Get-CimInstance Win32_BIOS          -ErrorAction SilentlyContinue
	$summaryLines.Add("OS            : $($os.Caption) Build $($os.BuildNumber)")
	$summaryLines.Add("Model         : $($cs.Manufacturer) $($cs.Model)")
	$summaryLines.Add("BIOS          : $($bios.SMBIOSBIOSVersion)")

	# Secure Boot CA state
	$sbReg = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing"
	$uefiStatus = $sbReg.UEFICA2023Status
	$uefiError  = $sbReg.UEFICA2023Error
	$conflevel  = $sbReg.ConfidenceLevel
	$sbEnabled  = Confirm-SecureBootUEFI -ErrorAction SilentlyContinue
	$summaryLines.Add("SecureBoot    : $sbEnabled")
	if ($uefierror) {
		$summaryLines.Add("UEFICA2023    : $uefiStatus [ConfidenceLevel: $conflevel] [ERROR: $uefiError]")
	} else {
		$summaryLines.Add("UEFICA2023    : $uefiStatus [ConfidenceLevel: $conflevel]")
	}
	$summaryLines.Add("")

	# Channel stats
	$summaryLines.Add("-" * 70)
	$summaryLines.Add("  Events Per Channel")
	$summaryLines.Add("-" * 70)
	foreach ($stat in $channelStats | Sort-Object EventCount -Descending) {
		if ($stat.EventCount -ge 0) {
			$summaryLines.Add("  $($stat.EventCount.ToString().PadLeft(5))  $($stat.Channel)")
		} else {
			$summaryLines.Add("  [N/A]  $($stat.Channel)  `(channel unavailable`)")
		}
	}
	$summaryLines.Add("")

	# High-interest events
	$summaryLines.Add("-" * 70)
	$summaryLines.Add("  HIGH-INTEREST EVENTS `(patterns relevant to 0xC1900101-0x20017`)")
	$summaryLines.Add("-" * 70)
	if (($flaggedEvents | Measure).Count -eq 0) {
		$summaryLines.Add("  (none found)")
	} else {
		foreach ($ev in $flaggedEvents | Sort-Object TimeCreatedUTC) {
			$summaryLines.Add("")
			$summaryLines.Add("  Time   : $($ev.TimeCreatedUTC) UTC")
			$summaryLines.Add("  Log    : $($ev.LogName)")
			$summaryLines.Add("  Level  : $($ev.Level)")
			$summaryLines.Add("  Id     : $($ev.EventId)")
			$summaryLines.Add("  Source : $($ev.ProviderName)")
			$summaryLines.Add("  Msg    : $($ev.Message.Substring(0, [Math]::Min(400, $ev.Message.Length)))")
		}
	}

	$summaryLines.Add("")
	$summaryLines.Add("-" * 70)
	$summaryLines.Add("  Output Files")
	$summaryLines.Add("-" * 70)
	Get-ChildItem $OutputDir -ErrorAction SilentlyContinue |
		Sort-Object LastWriteTime |
		ForEach-Object { $summaryLines.Add("  $($_.Name)  [$([Math]::Round($_.Length/1KB, 1)) KB]") }

	$summaryLines.Add("")
	$summaryLines.Add("=" * 70)

	$summaryLines | Out-File -FilePath $summaryPath -Encoding UTF8 -Force
	Write-Log "Summary written: $summaryPath"

	# Echo summary to stdout so it appears in smsts.log
	$summaryLines | ForEach-Object { Write-Output $_ }
}

Write-Log "=== Save-UpgradeEventLogs.ps1 completed. Total events: $(($allEvents | Measure).Count) ==="
exit 0
