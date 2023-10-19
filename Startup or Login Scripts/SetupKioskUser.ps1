# Creates user account if it doesn't exist, generating with a blank password.
# This should be combined with policies that set the user to automatically login.
# MJC 5-5-23
$KIOSK_USER = 'kioskuser0'
if ((Get-LocalUser | where {$_.name -eq $KIOSK_USER} | Measure-Object).Count -le 0) {
	New-LocalUser -Name $KIOSK_USER -NoPassword -Description "Kiosk Mode User Account (Auto-Login)" -AccountNeverExpires -UserMayNotChangePassword | Set-LocalUser -PasswordNeverExpires $true
	# Add to Users group.
	Add-LocalGroupMember -SID S-1-5-32-545 -Member $KIOSK_USER
}
