# Outputs all shared service account mailboxes owned or co-owned by the given user.
# If you want information on a single shared service account, use Find-ADEmailInfo.ps1 instead.
# MJC 7-10-23
$user=Read-Host "Enter Username"
$dn=(Get-ADUser $user).DistinguishedName
Get-ADGroup -LDAPFilter "(|(msExchCoManagedByLink:=$dn)(managedby=$dn))" -Properties msExchCoManagedByLink,ManagedBy,Member | where {$_.Name -like "grp-*" } | Select @{N="OLGroup"; Expression={$_.Name}}, @{N="Mailbox"; Expression={ if ($_.Name -match "grp\-(\w+)" -And ($mail = (Get-ADUser $matches[1] -Properties mail).mail)) { $mail } else { "<UNKNOWN MAILBOX>"}}},ManagedBy, msExchCoManagedByLink, Member, DistinguishedName | Out-GridView

Read-Host "Press enter to exit" | Out-Null

