# Rename a computer using smartcard credentials.
<#
    .SYNOPSIS
    Renames a remote computer using modern credentials.
    
	.DESCRIPTION
    Renames a remote computer using modern credentials (including virtual smartcards).
    
	.PARAMETER ComputerName
    Required. The remote computer name.
	
	.PARAMETER NewName
    Required. The new name for the computer.
	
	.PARAMETER Restart
	Restarts the computer after renaming.
	
	.EXAMPLE
	Rename-Computer-SC.ps1 -ComputerName "Foobar" -NewName "Dingle"
	
    .NOTES
    Author: MJC 10-12-2022
#>
<#
param(
	[Parameter(Mandatory=$true)]
	[string]$ComputerName,
	
	[Parameter(Mandatory=$true)]
	[string]$NewName, 
	
	[switch]$Restart
)
#>

# Copyright: (c) 2021, Jordan Borean (@jborean93) <jborean93@gmail.com>
# MIT License (see LICENSE or https://opensource.org/licenses/MIT)
# Source: https://gist.github.com/jborean93/586382429f869bdc1415d0ccb90db2e7
Function Get-ModernCredential {
    <#
    .SYNOPSIS
    Modern credential prompt.
    .DESCRIPTION
    Uses the modern Windows credential prompt to build a credential object.
    .PARAMETER Message
    The message to display in the credential prompt. Defaults to 'Enter your credentials.'.
    .PARAMETER Title
    The title to display in the credential prompt. Defaults to 'PowerShell credential request'.
    .PARAMETER Username
    Optional default user to prefill in the credential prompt. Can be combined with '-ForceUsername' to force this
    username in the credential prompt.
    .PARAMETER Win32Error
    The Win32 error code as an integer or Win32Exception that can be used to automatically display an error message in
    the credential prompt. By default (0) means do not display any error message.
    .PARAMETER ForceUsername
    Do not allow the caller to type in another username, must be set with '-Username'.
    .PARAMETER ShowCurrentUser
    Add the current user to the More choices pick list.
    .EXAMPLE Get a credential supplied by the user
    $cred = Get-ModernCredential
    .EXAMPLE Get the credential for 'username'
    $cred = Get-ModernCredential -Username username
    .EXAMPLE Get the credential for only 'username'
    $cred = Get-ModernCredential -Username username -ForceUsername
    .EXAMPLE Display error message from previous credential attempt
    # 5 -eq ERROR_ACCESS_DENIED
    $cred = Get-ModernCredential -Win32Error 5
    .NOTES
    This only works on Windows for both Windows PowerShell and PowerShell.
    #>
    [CmdletBinding()]
    [OutputType([PSCredential])]
    param (
        [Parameter()]
        [String]
        $Message = 'Enter your credentials.',

        [Parameter()]
        [String]
        $Title = 'PowerShell credential request',

        [Parameter()]
        [AllowEmptyString()]
        [String]
        $Username,

        [Parameter()]
        [Object]
        $Win32Error = 0,

        [Switch]
        $ForceUsername,

        [Switch]
        $ShowCurrentUser
    )

    begin {
        $addParams = @{}
        $addTypeCommand = Get-Command -Name Add-Type

        # CompilerParameters is used for Windows PowerShell only.
        if ('CompilerParameters' -in $addTypeCommand.Parameters.Keys) {
            $addParams.CompilerParameters = [CodeDom.Compiler.CompilerParameters]@{
                CompilerOptions = '/unsafe'
            }
        }
        else {
            $addParams.CompilerOptions = '/unsafe'
        }

        Add-Type @addParams -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
namespace ModernPrompt
{
    public class NativeHelpers
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public class CREDUI_INFO
        {
            public Int32 cbSize;
            public IntPtr hwndParent;
            public string pszMessageText;
            public string pszCaptionText;
            public IntPtr hbmBanner;
            public CREDUI_INFO()
            {
                this.cbSize = Marshal.SizeOf(this);
            }
        }
    }
    public class NativeMethods
    {
        [DllImport("credui.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredPackAuthenticationBuffer(
            Int32 dwFlags,
            string pszUserName,
            string pszPassword,
            IntPtr pPackedCredentials,
            ref Int32 pcbPackedCredentials);
        [DllImport("credui.dll", CharSet = CharSet.Unicode)]
        public static extern Int32 CredUIPromptForWindowsCredentials(
            NativeHelpers.CREDUI_INFO pUiInfo,
            Int32 dwAuthError,
            ref uint pulAuthPackage,
            IntPtr pvInAuthBuffer,
            uint ulInAuthBufferSize,
            out IntPtr ppvOutAuthBuffer,
            out uint pulOutAuthBufferSize,
            ref bool pfSave,
            Int32 dwFlags);
        [DllImport("credui.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool CredUnPackAuthenticationBuffer(
            Int32 dwFlags,
            IntPtr pAuthBuffer,
            uint cbAuthBuffer,
            StringBuilder pszUserName,
            ref Int32 pcchMaxUserName,
            StringBuilder pszDomainName,
            ref Int32 pcchMaxDomainame,
            IntPtr pszPassword,
            ref Int32 pcchMaxPassword);
        [DllImport("Ole32.dll")]
        public static extern void CoTaskMemFree(
            IntPtr pv);
        public static SecureString PtrToSecureStringUni(IntPtr buffer, int length)
        {
            unsafe
            {
                char *charPtr = (char *)buffer.ToPointer();
                return new SecureString(charPtr, length);
            }
        }
    }
}
'@

        $credUI = [ModernPrompt.NativeHelpers+CREDUI_INFO]@{
            pszMessageText = $Message
            pszCaptionText = $Title
        }

        $ERROR_INSUFFICIENT_BUFFER = 0x0000007A
        $ERROR_CANCELLED = 0x00004C7

        if ($Win32Error) {
            if ($Win32Error -is [ComponentModel.Win32Exception]) {
                $Win32Error = $Win32Error.NativeErrorCode
            }
        }
    }

    end {
        $inCredBufferSize = 0
        $inCredBuffer = [IntPtr]::Zero
        $outCredBufferSize = 0
        $outCredBuffer = [IntPtr]::Zero

        try {
            # If a default username is specified we need to specify an in credential buffer with that name
            if (-not [String]::IsNullOrWhiteSpace($Username)) {
                while ($true) {
                    $res = [ModernPrompt.NativeMethods]::CredPackAuthenticationBuffer(
                        0,
                        $Username,
                        '',
                        $inCredBuffer,
                        [ref]$inCredBufferSize
                    ); $err = [Runtime.InteropServices.Marshal]::GetLastWin32Error()

                    if ($res) {
                        break
                    }
                    elseif ($err -eq $ERROR_INSUFFICIENT_BUFFER) {
                        $inCredBuffer = [Runtime.InteropServices.Marshal]::AllocHGlobal($inCredBufferSize)
                    }
                    else {
                        $exp = [ComponentModel.Win32Exception]$err
                        Write-Error -Message "Failed to pack input username: $($exp.Message)" -Exception $exp
                        return
                    }
                }
            }

            $authPackage = 0
            $save = $false
            $flags = 0

            if ($ForceUsername) {
                $flags = $flags -bor 0x20  # CREDUIWIN_IN_CRED_ONLY
            }

            if ($ShowCurrentUser) {
                $flags = $flags -bor 0x200  # CREDUIWIN_ENUMERATE_CURRENT_USER
            }
    
            $err = [ModernPrompt.NativeMethods]::CredUIPromptForWindowsCredentials(
                $credUI,
                $Win32Error,
                [ref]$authPackage,
                $inCredBuffer,
                $inCredBufferSize,
                [ref]$outCredBuffer,
                [ref]$outCredBufferSize,
                [ref]$save,
                $flags
            )
    
            if ($err -eq $ERROR_CANCELLED) {
                return  # No credential was specified
            }
            elseif ($err) {
                $exp = [ComponentModel.Win32Exception]$err
                Write-Error -Message "Failed to prompt for credential: $($exp.Message)" -Exception $exp
                return
            }

            $usernameLength = 0
            $domainLength = 0
            $passwordLength = 0
            $usernameBuffer = [Text.StringBuilder]::new(0)
            $domainBuffer = [Text.StringBuilder]::new(0)
            $passwordPtr = [IntPtr]::Zero

            try {
                while ($true) {
                    $res = [ModernPrompt.NativeMethods]::CredUnpackAuthenticationBuffer(
                        1,  # CRED_PACK_PROTECTED_CREDENTIALS
                        $outCredBuffer,
                        $outCredBufferSize,
                        $usernameBuffer,
                        [ref]$usernameLength,
                        $domainBuffer,
                        [ref]$domainLength,
                        $passwordPtr,
                        [ref]$passwordLength
                    ); $err = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
    
                    if ($res) {
                        break
                    }
                    elseif ($err -eq $ERROR_INSUFFICIENT_BUFFER) {
                        [void]$usernameBuffer.EnsureCapacity($usernameLength)
                        [void]$domainBuffer.EnsureCapacity($passwordLength)
                        $passwordPtr = [Runtime.InteropServices.Marshal]::AllocHGlobal($passwordLength)
                    }
                    else {
                        $exp = [ComponentModel.Win32Exception]$err
                        Write-Error -Message "Failed to unpack credential: $($exp.Message)" -Exception $exp
                        return
                    }
                }

                # We want to avoid reading the password as a full string so use this "unsafe" method
                $password = [ModernPrompt.NativeMethods]::PtrToSecureStringUni($passwordPtr, $passwordLength)
            }
            finally {
                if ($passwordPtr -ne [IntPtr]::Zero) {
                    $blanks = [byte[]]::new($passwordLength * 2)  # Char takes 2 bytes
                    [Runtime.InteropServices.Marshal]::Copy($blanks, 0, $passwordPtr, $blanks.Length)
                    [Runtime.InteropServices.Marshal]::FreeHGlobal($passwordPtr)
                }
            }

            if ($domainLength) {
                $credUsername = '{0}\{1}' -f ($domainBuffer.ToString(), $usernameBuffer.ToString())
            }
            else {
                $credUsername = $usernameBuffer.ToString()
            }
            [PSCredential]::new($credUsername, $password)
        }
        finally {
            if ($outCredBuffer -ne [IntPtr]::Zero) {
                # Should be calling SecureZeroMemory but we cannot access this in .NET so do the next best thing
                # and wipe the unmanaged memory ourselves.
                $blanks = [byte[]]::new($outCredBufferSize)
                [Runtime.InteropServices.Marshal]::Copy($blanks, 0, $outCredBuffer, $blanks.Length)
                [ModernPrompt.NativeMethods]::CoTaskMemFree($outCredBuffer)
            }

            if ($inCredBuffer -ne [IntPtr]::Zero) {
                [Runtime.InteropServices.Marshal]::FreeHGlobal($inCredBuffer)
            }
        }
    }
}

# Prompt user if parameters not given on command-line.
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
if (-Not [string]::IsNullOrWhitespace($ComputerName)) {
	$_computerName = $ComputerName
} else {
	$_computerName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter a computer name:','ComputerName','')
}
if ([string]::IsNullOrWhitespace($_computerName)) {
	[Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: Invalid or blank computer name", "OKOnly,Critical", "Error") | Out-Null
} else {
	if (Test-Connection -ComputerName $_computerName -Count 1 -Quiet -ErrorAction Stop) {
		# Always prompt for credentials.
        try {
		    $cred = Get-ModernCredential -Win32Error 5
        } catch {
            $msg = $_.Exception.Message 
            [Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: $msg", "OKOnly,Critical", "Error") | Out-Null
            return -1
        }
		if ([string]::IsNullOrWhitespace($cred.Username)) {
			[Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: Invalid username", "OKOnly,Critical", "Error") | Out-Null
		} else {
			if (-Not [string]::IsNullOrWhitespace($NewName)) {
				$_newName = $NewName
			} else {
				$_newName = [Microsoft.VisualBasic.Interaction]::InputBox('Enter NEW computer name:','New ComputerName',$_computerName)
			}
			if ([string]::IsNullOrWhitespace($_newName)) {
				[Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: Invalid or blank computer name", "OKOnly,Critical", "Error") | Out-Null
			} else {
				if ($Restart -ne $null) {
					$AddParams = @{ "Restart"=$Restart }
				} else {
					$choice = [Microsoft.VisualBasic.Interaction]::MsgBox("Restart computer after rename?", "YesNo,Question", "Restart?")
					if ($choice -eq "Yes") {
						$AddParams = @{ "Restart"=$true }
					} else {
						$AddParams = @{}
					}
				}
				try {
                    Rename-Computer -ComputerName $_computerName -NewName $_newName -DomainCredential $cred -Force @AddParams
                } catch {
                    $msg = $_.Exception.Message 
                    [Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: $msg", "OKOnly,Critical", "Error") | Out-Null
                }
			}
		} 
	} else {
		[Microsoft.VisualBasic.Interaction]::MsgBox("ERROR: [${_computerName}] is offline or unreachable", "OKOnly,Critical", "Error") | Out-Null
	}
}
Read-Host "Press enter to exit" | Out-Null