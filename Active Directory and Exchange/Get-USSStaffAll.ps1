# Returns all staff in USS.

$outCSV = "{$ENV:OneDrive}\uss_staff.csv"
# Additionally search these groups.
$dynamicStaffGroups = "HW-HSA-ALL","AllUSS"
# Only output these employee types.
$filterEmployeeTypes = "Staff","Contractor"
# Results based on company field in AD.
$alluss = Get-ADUser -LDAPFilter "(company=JHU; University Student Services)" -Properties Department,Company,DisplayName,mail,distinguishedname,extensionattribute2
# Results based on dynamic group, to capture users in different divisions (like Ranstaad).
$morestaff = $dynamicStaffGroups | foreach { Get-ADGroup $_ -Properties Member | Select -ExpandProperty Member | where {$_ -notin $alluss.distinguishedname} | foreach { Get-ADUser $_ -Properties Department,Company,DisplayName,mail,distinguishedname,extensionattribute2} } | Sort -Unique distinguishedname
if ($alluss -isnot [array]) {
	$alluss = @($alluss)
}
if ($morestaff -isnot [array]) {
	$morestaff = @($morestaff)
}
$alluss += $morestaff
$alluss | where {$_.extensionattribute2 -in $filterEmployeeTypes -Or $filterEmployeeTypes -eq $null} | Select @{N="JHED"; Expression={$_.Name}}, mail, DisplayName, @{N="EmployeeType"; Expression={$_.extensionAttribute2}},Department,Company | Sort -Property "Department" | Export-CSV -NoTypeInformation $outCSV
