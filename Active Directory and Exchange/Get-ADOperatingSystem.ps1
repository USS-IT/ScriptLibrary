# Little snippets for getting real-time OS info remotely queried through WMI. System must be on the Hopkins network.
# Account used must have either local admin or remote admin rights on system.
# Requires RSAT AD Tools
# MJC 5-16-24

$comp = Read-Host "Enter Computer Name"
Get-ADComputer $comp -Properties OperatingSystemVersion
Read-Host "Press enter to exit" | Out-Null
