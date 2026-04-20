<#
	.SYNOPSIS
	Checks for network connectivity by pinging given host, displaying a pop-up if no connection found.
	
	.DESCRIPTION
	Checks for network connectivity by pinging given host, displaying a pop-up if no connection found.

  .PARAMETER Host
  Address to use to verify internet activity (default: "google.com").

  .PARAMETER ErrorOnPacketLoss
  Return an error code on ping failure.
  
  .PARAMETER SkipPopup
  Don't show a pop-up on ping failure.
	
  .NOTES
	Author: mcarras8
#>   
param(
   [Parameter(Mandatory=$false)]
   [ValidateNotNullorEmpty()]
   [string]$Host="google.com",

   [Parameter(Mandatory=$false)]
   [switch]$ErrorOnPacketLoss,

   [Parameter(Mandatory=$false)]
   [switch]$SkipPopup
)
$exitCode = 0
if (-Not ((Test-Connection -ComputerName $Host -Count 1 -Quiet) -Or 
		  (Test-Connection -ComputerName $Host -Count 1 -Quiet) -Or 
		  (Test-Connection -ComputerName $Host -Count 1 -Quiet))) {
  if (-Not $SkipPopup) {
	    $wshell = New-Object -ComObject Wscript.Shell
	    $wshell.Popup("Cannot reach [$Host]. System may be missing MAC registration.",0,"Warning - Address Unreachable",16+4096)
  }
  if ($ErrorOnPacketLoss) {
     $exitCode = -1
  }
}
[System.Environment]::Exit($exitCode)