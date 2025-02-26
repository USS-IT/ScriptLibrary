<#
	.SYNOPSIS
	Remotely disables or enables a local account on a given machine.
	
	.DESCRIPTION
	Remotely disables or enables a local account on a given machine. This must be run under an account with local admin access on the target computer.
	
	.NOTES
	Alternative method (also requires local admin):
	1) Open Computer Management
	2) Right-click and select "Connect to another computer..."
	3) Access Local users and groups.
	
	Author: mcarras8
#>
$_FLAGS_ACCOUNTDISABLE=0x0002

$ComputerName = Read-Host "Enter computer name"
If((Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) -Or (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet) -Or (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
	Write-Host "[$ComputerName] appears to be online, querying..."
} else {
	Write-Warning "[$ComputerName] did not respond to any ping attempts, querying anyway..."
}
$Username = Read-Host "Enter AD username"

try {
	$user = [ADSI]"WinNT://$ComputerName/$Username,User"
	
	Write-Host "Default is 'Y' (Disable)"
	$doDisable = "Disable account [$Username]? (Y/N)"
	if ($doDisable -ne "N") {
		# Disable account
		$newflags = $user.UserFlags.Value -bor $_FLAGS_ACCOUNTDISABLE
	} else {
		# Enable account
		$newflags = $user.UserFlags.Value -bxor $_FLAGS_ACCOUNTDISABLE
	}
	$user.put("userflags",$newflags)
	$user.SetInfo()
	# $user.refreshcache()
} catch {
	Write-Error $_
}
Read-Host "Press enter to exit" | Out-Null
