<#
	.SYNOPSIS
	Sets the "XInstallWebView2" task sequence variable depending on if the given version is older than the installed version of the WebView2 Runtime.
	
	.DESCRIPTION
	Sets the "XInstallWebView2" task sequence variable depending on if the given version is older than the installed version of the WebView2 Runtime.
	
	.PARAMETER TBIVersion
	The version to be installed as a string, e.g. "122.0.2365.106".
	
	.NOTES
	Author: Matt Carras (mcarras8)
#>
Param(
    [Parameter(Mandatory=$true)]
	[version]$TBIVersion
)
	
$RegPaths = @(
    "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients"
)

$InstalledVersion = $null
foreach ($Path in $RegPaths) {
    if (Test-Path $Path) {
        Get-ChildItem $Path | ForEach-Object {
            try {
                $pv = (Get-ItemProperty $_.PsPath -ErrorAction Stop).pv
                if ($pv -and (-not $InstalledVersion -or [version]$pv -gt $InstalledVersion)) {
                    $InstalledVersion = [version]$pv
                }
            } catch {}
        }
    }
}

$tsenv = New-Object -ComObject Microsoft.SMS.TSEnvironment
if (-not $InstalledVersion -or $InstalledVersion -lt $TBIVersion) {
    $tsenv.Value("XInstallWebView2") = "true"
} else {
    $tsenv.Value("XInstallWebView2") = "false"
}

[System.Environment]::Exit(0)