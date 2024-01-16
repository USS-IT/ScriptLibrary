<#
Script designed to scan an OU for computer objects that are inactive based on network connectivity and their last logon date. Inactive (stale) computers will be moved to a "retirement" OU, where they wille eventually be disabled, then deleted from AD. Results are recorded in a dated CSV on the HSA network share.
Created by Daniel Anderson - dander83@jhu.edu 
#>

#Import AD module for earlier versions of PowerShell
Import-Module ActiveDirectory

#Set time information. Computers that have not been logged into for 90 days are eligible to be retired. Computers that have not been logged into after 210 days are eligible to be delted from AD.
#Change the value after AddDays to customize the timeframes
$retirement = (Get-Date).AddDays(-90)
$removal = (Get-Date).AddDays(-210)
$CurrentDate = ((Get-Date).ToString('MM-dd-yyyy'))

#Configure CSV layout and name/location
$csvformat = @"
ComputerName,LastLogonDate,PingResult,Action
"@
$csvformat | Set-Content \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv

#Scan USS OU (SearchBase) for computers that have not been logged into in over 90 days (lastlogondate "less than" 90 days ago). Filters out Retiring OU, as the next section addresses that.
#Get-ADComputer results pinged. If a computer is on the network, it will do nothing to the AD Object and update the CSV accordingly. If it was not reachable, it will be moved to the retiring OU and CSV marked accordingly.
Get-ADComputer -Property Name,lastLogonDate -Filter {lastlogondate -lt $retirement} -SearchBase 'OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu' | Select-Object Name,lastLogonDate,DistinguishedName | Where-Object {$_.DistinguishedName -notlike "CN=*,OU=USS-Retired,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu"} 
|Where-Object {$_.DistinguishedName -notlike "CN=*OU=DMC-Gaming Loft,OU=DMC-Patron,OU=USS-DMC,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu"} |
ForEach-Object {
if (Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue) {
  Add-Content \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv "$($_.name),$($_.LastLogonDate),Success,Kept"
}Else {
  Add-Content \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv "$($_.name),$($_.LastLogonDate),Fail,Moved"
  Get-ADComputer $_.name | Move-ADObject -TargetPath 'OU=USS-Retired,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu'
}
}

#After the HSA OUs are scanned and any stale PCs moved to the retiring OU, the script will scan the Retiring OU to detect any computers that should not be in there, and disable and delete old objects
#If a computer can be pinged, the CSV will be updated and the LAN Admin should investigate.
#If a computer is not reachable but it has been logged into within 90 days, the CSV will be updated and the LAN Admin should investigate.
#If a computer has not been logged into for 365 days (the $removal date), AND the lastlogondate field is not null, it will be deleted from AD
#If a computer has not been logged into for 90 days and it is still enabled in AD, it will be disabled.
#If a computer is already disabled, it will be marked as "disabled" on the CSV
#Anything that does not fit the criteria above will be marked on the CSV as "Unknown"
Get-ADComputer -Property Name,lastLogonDate,DistinguishedName,Enabled -Filter * -SearchBase 'OU=USS-Retired,OU=Computers,OU=USS,DC=win,DC=ad,DC=jhu,DC=edu' | Select-Object Name,lastLogonDate,DistinguishedName,Enabled |
ForEach-Object {
$LogonNullCheck = $_.lastLogonDate
if (Test-Connection $_.name -Count 1 -ErrorAction SilentlyContinue) {
  Add-Content \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv "$($_.name),$($_.LastLogonDate),Success,CHECK - Online "
}Elseif ($_.LastLogreonDate -gt $retirement) {
  Add-Content \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv "$($_.name),$($_.LastLogonDate),Fail,CHECK - Last logon"
}Elseif ($_.LastLogonDate -le $removal -AND $LogonNullCheck) {
  Get-ADComputer $_.DistinguishedName | Remove-ADObject -Confirm:$false
  Add-Content \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv "$($_.name),$($_.LastLogonDate),Fail,Deleted"}
  Elseif ($_.LastLogonDate -le $retirement -AND $_.Enabled -eq "True") {
  Add-Content \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv "$($_.name),$($_.LastLogonDate),Fail,Disabled"
  Get-ADComputer $_.DistinguishedName | Disable-ADAccount
}Elseif ($_.Enabled -eq "False")  {
  Add-Content \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv "$($_.name),$($_.LastLogonDate),Fail,Disabled already"
}Else {
  Add-Content \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv "$($_.name),$($_.LastLogonDate),Fail,Unknown"
}
}

Send-MailMessage -From 'HSA IT Services <hsaitservices@jhu.edu>' -To 'HSA IT Services <hsaitservices@jhu.edu>' -Subject 'Stale PC Results' -Body "Results of Stale PC script attached" -Attachments \\win.ad.jhu.edu\cloud\HSA$\ITServices\Scripts\Results\StalePCs-$CurrentDate.csv -Priority High -DeliveryNotificationOption OnSuccess, OnFailure -SmtpServer 'smtp.johnshopkins.edu'