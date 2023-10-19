# Disable local admin remotely for $ComputerName.
# Requires local admin access on machine for current account.
# Alternative method (also requires local admin):
# 1) Open Computer Management
# 2) Right-click and select "Connect to another computer..."
# 3) Access Local users and groups.

$ComputerName = "SomeComputer"
If((Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet) -Or (Test-Connection -ComputerName $comp -Count 1 -Quiet)) {
	Write-Host "[$comp] appears to be online, querying..."
} else {
	Write-Warning "[$comp] did not respond to any ping attempts, querying anyway..."
}
$Username = "SomeUsername"
$_FLAGS_ACCOUNTDISABLE=0x0002
$user = [ADSI]"WinNT://$ComputerName/$Username,User"
# Disable account
$newflags = $user.UserFlags.Value -bor $_FLAGS_ACCOUNTDISABLE
# Enable account
# $newflags = $user.UserFlags.Value -bxor $_FLAGS_ACCOUNTDISABLE
$user.put("userflags",$newflags)
$user.SetInfo()
# $user.refreshcache()
Read-Host "Press enter to exit" | Out-Null
