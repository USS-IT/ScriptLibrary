<#
    .SYNOPSIS
    This function has three ways of being used. It can be used to track logon history, logoff history, and/or can email logs to the specified address.

    .DESCRIPTION
    Each time the function is run, it will perform it's setup, which involves setting up the log path, and replicating the script file locally
    for the send email portion if used through a scheduled task. Once complete, depending on the action parameter, will act accordingly.

    .PARAMETER action
    Required. Either "Logon", "Logoff", or "Email" to specific what the script should do.

    .PARAMETER from
    Optional. From address can be specified if sending an email. Must be paired with "Email" as action to take effect.

    .PARAMETER to
    Optional. To address can be specified if sending an email. Must be paired with "Email" as action to take effect.

    .NOTES
    Version:        1.0
    Author:         LDG
    Creation Date:  01/02/2024
    Purpose/Change: Track logon/off histories on computers.
#>

#-----------------------------------------------------------[Parameters]-----------------------------------------------------------
param(
    [Parameter(Mandatory=$true, Position=0)]
    [String]$action,

    [mailaddress]$from=[mailaddress]'ussitcloudapps@jhu.edu',

    [mailaddress]$to=[mailaddress]'ussitservices@jh.edu'
)
#---------------------------------------------------------[Initialisations]--------------------------------------------------------

# Ensure best coding practices are followed
Set-StrictMode -Version Latest

# Set Error Action to Continue
$ErrorActionPreference = "Continue"

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$today=$(Get-Date)
$yesterday=$((Get-Date).AddDays(-1))
$folderPath="C:\Logs\USS-LogonHistory"
$scriptPath=$folderPath + "\logonHistory.ps1"
$todayLogPath=$folderPath + "\logonHistory - $(Get-Date -Date $today -UFormat "%m-%d-%Y") - $($env:computername).csv"
$yesterdayLogPath=$folderPath + "\logonHistory - $(Get-Date -Date $yesterday -UFormat "%m-%d-%Y") - $($env:computername).csv"
$smtpServer='smtp.johnshopkins.edu'

#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
    .DESCRIPTION
    Checks to see if the log path exists and creates it if it doesn't. Also downloads a local copy of this script to be used by the task scheduler.
    Additionally create today's logfile if it does not already exist.
 #>
function setupCheck {
    # If the log folder path does not exist, create it!
    If(-Not (Test-Path -Path $folderPath)) {
        New-Item -Path $folderPath -ItemType "directory"
    }

    # If the script has not been downloaded locally, download it! It's required for the the scheduled task
    # to send the logfile in an email.
    If(-Not (Test-Path -Path $scriptPath -PathType Leaf)) {
        $scriptLink = "https://raw.githubusercontent.com/USS-IT/ScriptLibrary/main/Startup%20or%20Login%20Scripts/LogonHistory.ps1"
        Invoke-WebRequest -URI $scriptLink -OutFile $scriptPath -UseBasicParsing
    }

    # If today's log file does not exist, create it!
    If (-Not (Test-Path -Path $todayLogPath -PathType Leaf)) {
        createTodayLogFile
    }
}

<#
    .DESCRIPTION
    Creates today's logfile. Assumes that one does not already exist.
 #>
function createTodayLogFile {
    New-Item -Path $todayLogPath -ItemType "file"
    Set-Content -Path $todayLogPath -Value "date,time,type,jhed"
}

<#
    .DESCRIPTION
    Deletes yesterdays's logfile.
 #>
function deleteYesterdayLogFile {
    Remove-Item -Path $yesterdayLogPath
}

<#
    .DESCRIPTION
    Based on the action parameter, this function will either log the current user as logged in, log the current user and logged out, or email yesterday's logfile
    and delete it once sent.
 #>
function handleAction {
    switch ($action)
    {
        "Logon" {
            # Log the current user and logged on in today's logfile.
            Add-Content -Path $todayLogPath -Value "$(Get-Date -Date $today -UFormat "%m/%d/%Y"),$(Get-Date -UFormat "%R"),login,$Env:UserName"
        }
        "Logoff" {
            # Log the current user and logged off in today's logfile.
            Add-Content -Path $todayLogPath -Value "$(Get-Date -Date $today -UFormat "%m/%d/%Y"),$(Get-Date -UFormat "%R"),logoff,$Env:UserName"
        }
        "Email" {
            # Email yesterday's logfile to the intended address.
            Send-MailMessage -From $from -To $to -Subject "$($env:computername): $(Get-Date -Date $yesterday -UFormat "%m/%d/%Y") Logs" -SmtpServer $smtpServer -Attachments $yesterdayLogPath

            # Now that we've sent the log history, delete yesterday's logfile.
            deleteYesterdayLogFile
        }
        Default {
            # do nothing.
        }
    }
}

<#
    .DESCRIPTION
    The main function calls all other functions to perform the actions required by this script.
#>
function main {
    setupCheck
    handleAction
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

# Call the main function
main

#-----------------------------------------------------------[Finish up]------------------------------------------------------------