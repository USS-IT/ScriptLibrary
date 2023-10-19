# Disable and re-enable WSUS listing in Policy to install RSAT tools.
# Remove the comment on the line for the Group Policy Management Tools if needed.

# Temporarily disable Windows Update policy, restarting Windows Update.
$UseWUServer = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" | Select-Object -ExpandProperty UseWUServer
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 0
Restart-Service "Windows Update"

# Install RSAT AD Management Tools
Get-WindowsCapability -Name "Rsat.ActiveDirectory.DS-LDS.Tools*" -Online | Add-WindowsCapability -Online

# Uncomment line below to also install RSAT Group Policy Management tools, if needed.
#Get-WindowsCapability -Name "Rsat.GroupPolicy.Management.Tools*" -Online | Add-WindowsCapability -Online

# Restore Windows Update policy and restart Windows Update service.
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value $UseWUServer
Restart-Service "Windows Update"
