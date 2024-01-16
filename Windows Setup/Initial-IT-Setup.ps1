<#
    .SYNOPSIS
    Installs base applications that members of the USS IT Services department use.

    .DESCRIPTION
    This script will install RSAT AD and Group Policy Management Tools, Configuration Manager Console, and a few other applications via winget.

    .NOTES
    Version:        1.0
    Author:         LDG
    Creation Date:  10/19/2023
    Purpose/Change: Automate setup process for USS IT admin systems.
#>

#-----------------------------------------------------------[Parameters]-----------------------------------------------------------

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Ensure best coding practices are followed
Set-StrictMode -Version Latest

# Set Error Action to Continue
$ErrorActionPreference = "Continue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

# Declare log types and their values.
$normalLog = "Normal"
$warningLog = "Warning"
$errorLog = "Error"

#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
    .DESCRIPTION
    The installRsatTools function installs RSAT AD and Group Policy Management Tools. During this process it will disable and re-enable WSUS listing in Policy to prevent Windows Updates.
 #>
function installRsatTools {
    # Temporarily disable Windows Update policy, restarting Windows Update.
    printLog $normalLog $(getLogInfo) "Disabling Windows Update..."
    $UseWUServer = Get-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" | Select-Object -ExpandProperty UseWUServer
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value 0
    Restart-Service "Windows Update"

    # Install RSAT AD Management Tools
    printLog $normalLog $(getLogInfo) "Installing Rsat AD tools..."
    Get-WindowsCapability -Name "Rsat.ActiveDirectory.DS-LDS.Tools*" -Online | Add-WindowsCapability -Online
    printLog $normalLog $(getLogInfo) "Installed Rsat AD tools"

    # Install RSAT Group Policy Management Tools
    printLog $normalLog $(getLogInfo) "Installing Rsat GP tools..."
    Get-WindowsCapability -Name "Rsat.GroupPolicy.Management.Tools*" -Online | Add-WindowsCapability -Online
    printLog $normalLog $(getLogInfo) "Installed Rsat GP tools"

    # Restore Windows Update policy and restart Windows Update service.
    printLog $normalLog $(getLogInfo) "Enabling Windows Update..."
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWUServer" -Value $UseWUServer
    Restart-Service "Windows Update"
}

function installConfigurationManagerConsole {
    $newConfigurationManagerFilePath = "\\win.ad.jhu.edu\data\emmsstuff$\System Center 2012\SCCM 2012 R2\SCCM 2107 Console Setup\"
    $tempDriveName = "SCCMInstall-Temp"
    New-PSDrive -Name $tempDriveName -PSProvider "FileSystem" -Root $newConfigurationManagerFilePath

    printLog $normalLog $(getLogInfo) "Running SCCM installer..."
    Invoke-Expression $($tempDriveName + ":\ConsoleSetup.exe TargetDir='C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\' DefaultSiteServerName=JHECMCAS.win.ad.jhu.edu EnableSQM=0")

    # todo: update statement to install silently. Currently fails to install if /q is added. Maybe need to retry using Start-Process.
}

<#
    .DESCRIPTION
    The installLatestWinget function installs the latest winget version.
    Spcial thanks to Michael Herrmann for original function: https://winget.pro/winget-install-powershell/
 #>
function installLatestWinget {
    # Get information about the latest winget installer from GitHub:
    $apiUrl = "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
    $apiInfo = $(Invoke-RestMethod $apiUrl)
    $wingetLatestVersion = $apiInfo.tag_name
    $wingetCurrentVersion = $(winget -v)

    if($wingetLatestVersion -ne $wingetCurrentVersion) {
        $wingetDownloadLink = $apiInfo.assets.browser_download_url | Where-Object {$_.EndsWith(".msixbundle")}

        # Download the installer:
        printLog $normalLog $(getLogInfo) "Downloading latest Winget version..."
        Invoke-WebRequest -URI $wingetDownloadLink -OutFile winget.msixbundle -UseBasicParsing

        # Install winget:
        printLog $normalLog $(getLogInfo) "Installing latest Winget version..."
        Add-AppxPackage winget.msixbundle
        printLog $normalLog $(getLogInfo) "Installed latest Winget version"

        # Remove the installer:
        Remove-Item winget.msixbundle

        printLog $normalLog $(getLogInfo) "Installed latest Winget version"
    } else {
        printLog $normalLog $(getLogInfo) "Latest Winget version is already installed!"
    }
}

<#
    .DESCRIPTION
    The installWingetApplications function installs all applications whose id's are listed in $wingetAppIds.
#>
function installWingetApplications {

    # Declare an array of all Winget application ids.
    $wingetAppIds = @(
        "Yubico.Authenticator",
        "Yubico.YubikeyManager",
        "Notepad++.Notepad++",
        "DominikReichl.KeePass",
        "Git.Git",
        "Microsoft.WindowsTerminal",
        "Microsoft.SQLServerManagementStudio",
        "Microsoft.PowerToys",
        "Dell.CommandUpdate.Universal",
        "Mozilla.Firefox",
        "VideoLAN.VLC"
    )
    # Declare an array of all Microsoft Store application ids.
    $msstoreAppIds = @(
    )

    $additionalAppsResponse = Read-Host "Would you like to specific additional applications to install via winget (please see readme for formatting)? (y/n)"

    if($additionalAppsResponse.Equals("y")){
        $additionalAppsInfo = Read-Host "Please enter a url or a filepath"

        # If information provided is a local path, grab the local json. Otherwise, do a web request.
        if (Test-Path -Path $additionalAppsInfo -PathType leaf) {
            $appsJson = Get-Content $additionalAppsInfo -Raw | ConvertFrom-Json
        } else {
            $appsJson = (Invoke-WebRequest -URI $additionalAppsInfo).Content | ConvertFrom-Json
        }

        foreach ($app in $appsJson.wingetApps){
            $wingetAppIds += $app.id
        }

        foreach ($app in $appsJson.msstoreApps){
            $msstoreAppIds += $app.id
        }
    }

    # Iterate over $wingetAppIds and install each one.
    foreach ($id in $wingetAppIds) {
        $wingetOutput = @(winget install -e --id $id --accept-package-agreements --accept-source-agreements)
        $message = ($id + " - " + $wingetOutput[-1])

        # Determine console log type based on successful install
        if ($message.Contains("Successfully installed")) {
            printLog $normalLog $(getLogInfo) $message
        } else {
            printLog $warningLog $(getLogInfo) $message
        }
    }

    # Iterate over $msstoreAppIds and install each one.
    foreach ($id in $msstoreAppIds) {
        $wingetOutput = @(winget install -e --id $id --source msstore --accept-package-agreements --accept-source-agreements)
        $message = ($id + " - " + $wingetOutput[-1])

        # Determine console log type based on successful install
        if ($message.Contains("Successfully installed")) {
            printLog $normalLog $(getLogInfo) $message
        } else {
            printLog $warningLog $(getLogInfo) $message
        }
    }

}

<#
    .DESCRIPTION
    setTaskbar overwrites the existing LayoutModification.xml file in the user's %LocalAppData%\Microsoft\Windows\Shell directory.
 #>
function setLayout {
    $username = $($Env:UserName)
    $filePath = "C:\Users\" + $username + "\AppData\Local\Microsoft\Windows\Shell\LayoutModification.xml"

    $personalLayoutResponse = Read-Host "Would you like to specify your own LayoutModification.xml file? (y/n)"

    if ($personalLayoutResponse.Equals("y")){
        $personalLayoutInfo = Read-Host "Please enter a url or a filepath"

        # If information provided is a local path, grab the local json. Otherwise, do a web request.
        if (Test-Path -Path $personalLayoutInfo -PathType leaf) {
            Copy-Item $personalLayoutInfo -Destination $filePath
        } else {
            Invoke-WebRequest -URI $personalLayoutInfo -OutFile $filePath -UseBasicParsing
        }
    } else {
        printLog $normalLog $(getLogInfo) "Setting default layout..."

        #todo: update me once merged with main!
        $newTaskbarLayoutLink = "https://raw.githubusercontent.com/USS-IT/ScriptLibrary/Initial-IT-Setup-Additional-Functionality/Windows%20Setup/LayoutModification.xml"
        Invoke-WebRequest -URI $newTaskbarLayoutLink -OutFile $filePath -UseBasicParsing
    }

    printLog $normalLog $(getLogInfo) "Layout updated! Please restart for changes to take effect."

}

<#
    .DESCRIPTION
    The printLog function takes in three parameters, the type of log, the function it originated from, and the log message and outputs it to the console.

    .PARAMETER type
    Specifies the type of output to write.

    .PARAMETER function
    Specifies the name of the function passing the log message.

    .PARAMETER message
    Specifies the message of the log.

#>
function printLog ([string]$type, [Array]$logInfo, [string]$message) {
    switch -Exact ($type) {
        $normalLog {Write-Output ($logInfo[0] + ": " + $message); Break}
        $warningLog {Write-Warning ($logInfo[0] + ": " + $message); Break}
        $errorLog {Write-Error ("Line: " + $logInfo[1] + ", " + $logInfo[1] + ": " + $message); Break}
        Default {Write-Error "Something went very wrong!"}
    }
}

<#
    .DESCRIPTION
    The getLogInfo function parses the call-stack to get information about the function that needs output a log.
 #>
function getLogInfo {
    $callStack = $(Get-PSCallStack)
    $name = $callStack.FunctionName[1]
    $line = $callStack.ScriptLineNumber[1]
    return $($name,$line)
}

<#
    .DESCRIPTION
    The main function calls all other functions to perform the actions required by this script.
#>
function main {
    installRsatTools
    installConfigurationManagerConsole
    installLatestWinget
    installWingetApplications
    setLayout
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Call the main function
main

#-----------------------------------------------------------[Finish up]------------------------------------------------------------
