
<h1 align="center">
  USS IT Script Library
</h1>

<h4 align="center">A repository of all frequently used scripts.</h4>

## How To Use

To clone and work in this repo, you'll need [Git](https://git-scm.com) installed on your computer. From your command line:

```bash
# Clone this repository
$ git clone https://github.com/USS-IT/ScriptLibrary.git

# Go into the repository
$ cd ScriptLibrary

# Make sure you create your own branch and get to work!
# You'll need to make a pull request for any changes to be pushed to production.
```

# Documentation
Click on the arrow by the script name to see brief documentation and usage notes.

## Active Directory and Exchange

<details>
<summary>Export-ADUsersToCSV.ps1</summary>
Output the members of a group to CSV file or show it in a pop-up. This includes nested members.
</details>

<details>
<summary>Find-ADEmailInfo.ps1</summary>
Search for accounts with primary mail and aliases using given wildcard *, showing information about any associated mail management (OLGroups) found.
	
For example, `*communications*` will show accounts with the word "communications" included in their primary or alias email addresses.
</details>

<details>
<summary>Find-ADMailboxesByOwner.ps1</summary>
Outputs all shared service account mailboxes owned or co-owned by the given user.
</details>

<details>
<summary>Get-ADEoLComputers.ps1</summary>
Outputs information on EOL computers from AD.
</details>

<details>
<summary>Get-ADOperatingSystem.ps1</summary>
Get OperatingSystemVersion reported by AD for given computer name. 
</details>

<details>
<summary>Get-ADUserGroups.ps1</summary>
Outputs a user's groups to a CSV file or show it in a pop-up. Can include all nested groups. 
</details>

<details>
<summary>Get-ADUserInfoFromList.ps1</summary>
Returns a list of users from AD given a CSV containing their Emails or JHEDs. Input CSV file must have column header "User". This can be either an email address (including aliases), UPN, or username/JHED.
</details>

<details>
<summary>Get-DMCGroupMembers.ps1</summary>
Show DMC group memberships for given domain user. Does NOT require RSAT tools.
</details>

<details>
<summary>Get-USSStaffAll.ps1</summary>
Outputs all staff and contractors in USS where company="USS" to a CSV file "uss_staff.csv" saved in OneDrive.
</details>

## Asset management

<details>
<summary>Create-AssetImportFiles.ps1</summary>
Creates import files for Snipe-It, SCCM, and JHARS from a Dell report exported as CSV. Can copy the files for SCCM and Snipe-It into their import paths.
</details>

<details>
<summary>Create-SCCMImportFromSnipeit.ps1</summary>
Creates an import file for SCCM from a previous Snipe-It export for re-importing a computer that's fallen out of SCCM, optionally copying over the file to the SCCM import path.
</details>

<details>
<summary>Get-StalePCAssetInfo.ps1</summary>
Compiles a report of assets to be deleted cross-referenced with our SOR (Snipe-It). The asset report must have the "ComputerName" column.
</details>

## Network Shares

<details>
<summary>Get-ACLReport.ps1</summary>
Returns all changed security permissions (ACL) and owners for subfolders starting at given root folder. Can compile results into a CSV file. This may take a while to complete when run on a remote network share.
</details>

## Printer Management

<details>
<summary>Get-PrinterInfo.ps1</summary>
Gets installed printer info for an online computer. Attempts to resolve WSD to IP addresses.
</details>

## Remote Administration

<details>
<summary>Enable-RemoteDesktop.ps1</summary>
Remotely enables Remote Desktop for the target computer using WMI.
</details>

<details>
<summary>Enable-RemoteWifiAdapter.ps1</summary>
Uses WMI to remotely enable any disabled WiFi adapters for a system online with ethernet.
</details>

<details>
<summary>Get-WMIBatteryHealth.ps1</summary>
Uses WMI to query info on an online computer's battery health and current capacity.
</details>

<details>
<summary>Get-WMIMemorySlots.ps1</summary>
Uses WMI to query info on an online computer's installed and free memory slots.
</details>

<details>
<summary>Get-WMIOperatingSystem.ps1</summary>
Uses WMI to query an online computer's operating system info.
</details>

<details>
<summary>Get-WMISoftware.ps1</summary>
Uses WMI to query an online computer's installed software. This should match Add/Remove Programs.
</details>

<details>
<summary>Get-WMIStorageInfo.ps1</summary>
Uses WMI to query an online computer's disk info (free space, type, etc.)
</details>

<details>
<summary>Get-WMIUptime.ps1</summary>
Uses WMI to query an online computer's uptime. Note Power > Shutdown will not reset this value.
</details>

<details>
<summary>Manage-LocalAdmin.ps1</summary>
Remotely disables or enables a local account on a given machine. This must be run under an account with local admin access on the target computer.
</details>

<details>
<summary>Rename-Computer-SC.ps1</summary>
Renames a remote computer using modern credentials (including virtual smartcards).
</details>

