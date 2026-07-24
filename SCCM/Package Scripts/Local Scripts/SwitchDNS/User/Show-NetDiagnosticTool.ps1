# Interactive network diagnostic tool for standard users
#
# - nslookup (with public DNS fallback)
# - ipconfig /flushdns
# - traceroute
#
# Author: Matt Carras (mcarras8)
# Created: 7-24-26

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ----------------------------
# Configuration
# ----------------------------

# Define selectable addresses in dropdown
$SelectableAddresses = @(
	"vpn.jh.edu",
    "jhu.edu",
    "johnshopkins.edu",
    "hopkinsmedicine.org",
	"login.jh.edu",
	"login.microsoftonline.com",
	"microsoft.com"
)

# Define additional public DNS to check DNS against.
$AddlPublicDNS = @(
	'8.8.8.8',
	'1.1.1.1'
)

# Title for the form and message boxes.
$UITitle = "IT Network Diagnostics"

$LogDir = "C:\USS\Logs\User\SwitchDNS"
# -- END Configuration --

# Create the log path if it doesn't already exist
if (-not (Test-Path $LogDir -PathType Container)) {
	$null = New-Item -Path $LogDir -ItemType Directory
}
try {
	$_scriptName = Split-Path -Leaf $PSCommandPath
} catch {
	$_scriptName = "Show-NetDiagnosticTool.ps1"
}
$LogPath = "$LogDir\$($_scriptName).log"
Start-Transcript $LogPath -Force

# Use WScript popup instead of MessageBox static calls
$script:Popup = New-Object -ComObject WScript.Shell

# Row number -> controls in that row
$script:ControlRows = @{}

# Create the main script form.
$FormMain = New-Object System.Windows.Forms.Form

# Add domain to SelectableAddresses, if it exists
$JoinedDomain = (Get-CimInstance Win32_ComputerSystem | Select -ExpandProperty Domain)
if (-Not [string]::IsNullOrEmpty($JoinedDomain)) {
	$SelectableAddresses += $JoinedDomain
}

# ----------------------------
# Helper functions
# ----------------------------

function Add-FormControl {
	# Add a new control to our main form.
    param(
		[Parameter(Mandatory=$true, Position=0)]
        [System.Windows.Forms.Control]$Control,
		
		[Parameter(Mandatory=$true, Position=1)]
        [int]$Row
    )

    if (-not $script:ControlRows.ContainsKey($Row)) {
        $script:ControlRows[$Row] = New-Object System.Collections.ArrayList
    }

	[void]$script:ControlRows[$Row].Add($Control)
	[void]$script:FormMain.Controls.Add($Control)
}

function Get-ControlRelativeLocation {
	# Get the relative Location point for the given row for Windows Forms.
    param(
		[Parameter(Mandatory=$true, Position=0)]
        [int]$Row,
				
        [int]$Spacing = 10,
		
        [int]$MarginLeft = 10,
		
        [int]$MarginTop = 10
    )

	$X = $MarginLeft
	$Y = $MarginTop
    # First control on first row
    if (-not $script:ControlRows.ContainsKey($Row) -Or $script:ControlRows[$Row].Count -eq 0) {		
		if ($script:ControlRows.ContainsKey($Row - 1)) {
			$PreviousRow = $script:ControlRows[$Row - 1]
			if ($PreviousRow.Count -gt 0) {
				$TallestBottom = (
					$PreviousRow |
					ForEach-Object { $_.Bottom } |
					Measure-Object -Maximum
				).Maximum

				$Y = $TallestBottom + $Spacing
			}
		}
    } else {
		$LastControl =
			$script:ControlRows[$Row][
				$script:ControlRows[$Row].Count - 1
			]
			
		$X = $LastControl.Right + $Spacing
		$Y = $LastControl.Top
	}

    return New-Object System.Drawing.Point($X, $Y)
}

# Function to get the target address (either from dropdown or custom input)
function Get-TargetAddress {
    if ($comboBoxAddress.SelectedItem -eq "-- Custom Address --") {
        $customAddr = $textBoxCustom.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($customAddr)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter a custom address (domain or IP).",
                "Input Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return $null
        }
        return $customAddr
    }
    else {
        return $comboBoxAddress.SelectedItem
    }
}

# Function to validate if input is an IP address
function Test-IsIPAddress {
    param ([string]$Address)
    
    $ipPattern = '^(\d{1,3}\.){3}\d{1,3}$'
    return $Address -match $ipPattern
}

# Function to enable/disable buttons during job execution
function Set-ButtonState {
    param([bool]$Running)
    
    $buttonDNS.Enabled = -not $Running
    $buttonTraceroute.Enabled = -not $Running
    $buttonFlushDNS.Enabled = -not $Running
    $buttonCancel.Enabled = $Running
    $buttonSaveLog.Enabled = -not $Running
    $comboBoxAddress.Enabled = -not $Running
    $textBoxCustom.Enabled = -not $Running
	$buttonAdapterInfo.Enabled = -not $Running
}

# -- END FUNCTIONS --

# -- START MAIN SCRIPT --

# Global variable to track running job and last operation
$script:currentJob = $null
$script:lastOperation = "DNS"  # Default operation type for filename

# Create the main form
$FormMain = New-Object System.Windows.Forms.Form
$FormMain.Text = $UITitle
$FormMain.Size = New-Object System.Drawing.Size(800, 600)
$FormMain.StartPosition = "CenterScreen"
$FormMain.FormBorderStyle = "FixedDialog"
$FormMain.MaximizeBox = $false

# Define margins and spacing
$marginLeft = 10
$marginTop = 10
$spacing = 10
$buttonHeight = 30
$labelHeight = 20

# Row 0
$row = 0
# Create address selection label
$labelSelectAddress = New-Object System.Windows.Forms.Label
$labelSelectAddress.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$labelSelectAddress.Size = New-Object System.Drawing.Size(150, $labelHeight)
$labelSelectAddress.Text = "Select Address to Check:"
Add-FormControl $labelSelectAddress $row

# Create address dropdown
$comboBoxAddress = New-Object System.Windows.Forms.ComboBox
$comboBoxAddress.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$comboBoxAddress.Size = New-Object System.Drawing.Size(200, $labelHeight)
$comboBoxAddress.DropDownStyle = "DropDownList"
foreach ($address in $SelectableAddresses) {
    $comboBoxAddress.Items.Add($address) | Out-Null
}
$comboBoxAddress.Items.Add("-- Custom Address --") | Out-Null
$comboBoxAddress.SelectedIndex = 0
Add-FormControl $comboBoxAddress $row

# Create custom address label
$labelCustom = New-Object System.Windows.Forms.Label
$labelCustom.Location = Get-ControlRelativeLocation -Row $row -Spacing 20 -MarginLeft $marginLeft -MarginTop $marginTop 
$labelCustom.Size = New-Object System.Drawing.Size(100, $labelHeight)
$labelCustom.Text = "Custom Address:"
$labelCustom.Visible = $false
Add-FormControl $labelCustom $row

# Create custom address textbox
$textBoxCustom = New-Object System.Windows.Forms.TextBox
$textBoxCustom.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$textBoxCustom.Size = New-Object System.Drawing.Size(($FormMain.ClientSize.Width - $labelCustom.Right - $spacing - $marginLeft), $labelHeight)
$textBoxCustom.Visible = $false
$textBoxCustom.Text = ""
Add-FormControl $textBoxCustom $row

# Row 1
$row++
# Create DNS Check button
$buttonDNS = New-Object System.Windows.Forms.Button
$buttonDNS.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$buttonDNS.Size = New-Object System.Drawing.Size(140, $buttonHeight)
$buttonDNS.Text = "Check DNS Servers"
$buttonDNS.BackColor = [System.Drawing.Color]::LightBlue
Add-FormControl $buttonDNS $row

# Create Traceroute button
$buttonTraceroute = New-Object System.Windows.Forms.Button
$buttonTraceroute.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$buttonTraceroute.Size = New-Object System.Drawing.Size(140, $buttonHeight)
$buttonTraceroute.Text = "Perform Traceroute"
$buttonTraceroute.BackColor = [System.Drawing.Color]::LightGreen
Add-FormControl $buttonTraceroute $row

# Create Flush DNS button
$buttonFlushDNS = New-Object System.Windows.Forms.Button
$buttonFlushDNS.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$buttonFlushDNS.Size = New-Object System.Drawing.Size(100, $buttonHeight)
$buttonFlushDNS.Text = "Flush DNS"
$buttonFlushDNS.BackColor = [System.Drawing.Color]::LightYellow
Add-FormControl $buttonFlushDNS $row

# Create Adapters button
$buttonAdapterInfo = New-Object System.Windows.Forms.Button
$buttonAdapterInfo.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$buttonAdapterInfo.Size = New-Object System.Drawing.Size(120, $buttonHeight)
$buttonAdapterInfo.Text = "Adapter Info"
$buttonAdapterInfo.BackColor = [System.Drawing.Color]::LightBlue
Add-FormControl $buttonAdapterInfo $row

# Row 3
$row++
# Create Save Log button 
$buttonSaveLog = New-Object System.Windows.Forms.Button
$buttonSaveLog.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$buttonSaveLog.Size = New-Object System.Drawing.Size(100, $buttonHeight)
$buttonSaveLog.Text = "Save Log"
$buttonSaveLog.BackColor = [System.Drawing.Color]::LightCyan
Add-FormControl $buttonSaveLog $row

# Create Clear button 
$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$buttonClear.Size = New-Object System.Drawing.Size(100, $buttonHeight)
$buttonClear.Text = "Clear Results"
$buttonClear.BackColor = [System.Drawing.Color]::LightCoral
Add-FormControl $buttonClear $row

# Create Cancel button 
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$buttonCancel.Size = New-Object System.Drawing.Size(100, $buttonHeight)
$buttonCancel.Text = "Cancel"
$buttonCancel.BackColor = [System.Drawing.Color]::Orange
$buttonCancel.Enabled = $false
Add-FormControl $buttonCancel $row

# Row 4
$row++
# Create results text box (fills remaining space)
$textBoxResults = New-Object System.Windows.Forms.TextBox
$textBoxResults.Location = Get-ControlRelativeLocation -Row $row -MarginLeft $marginLeft -MarginTop $marginTop -Spacing $spacing
$textBoxResults.Size = New-Object System.Drawing.Size(
    ($FormMain.ClientSize.Width - $textBoxResults.Location.X - (2 * $marginLeft)), 
	($FormMain.ClientSize.Height - $textBoxResults.Location.Y - $marginTop)
)
$textBoxResults.Multiline = $true
$textBoxResults.ScrollBars = "Vertical"
$textBoxResults.Font = New-Object System.Drawing.Font("Consolas", 9)
$textBoxResults.ReadOnly = $true
$textBoxResults.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                         [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                         [System.Windows.Forms.AnchorStyles]::Left -bor 
                         [System.Windows.Forms.AnchorStyles]::Right
Add-FormControl $textBoxResults $row

# Create a timer for updating job output
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 250  # Update every 250ms

# Event handler for dropdown selection change
$comboBoxAddress.Add_SelectedIndexChanged({
    if ($comboBoxAddress.SelectedItem -eq "-- Custom Address --") {
        $labelCustom.Visible = $true
        $textBoxCustom.Visible = $true
        $textBoxCustom.Focus()
    }
    else {
        $labelCustom.Visible = $false
        $textBoxCustom.Visible = $false
    }
})

# Timer tick event to check job status and update output
$timer.Add_Tick({
    if ($script:currentJob -ne $null) {
        # Check if job is still running
        if ($script:currentJob.State -eq 'Running') {
            # Receive any new output
            $output = Receive-Job -Job $script:currentJob
            if ($output) {
                $textBoxResults.AppendText($output)
            }
        }
        else {
            # Job completed or stopped
            $output = Receive-Job -Job $script:currentJob
            if ($output) {
                $textBoxResults.AppendText($output)
            }
            
            if ($script:currentJob.State -eq 'Completed') {
                $textBoxResults.AppendText("`r`n" + "=" * 80 + "`r`n")
                $textBoxResults.AppendText("Operation completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n")
            }
            elseif ($script:currentJob.State -eq 'Stopped') {
                $textBoxResults.AppendText("`r`n" + "=" * 80 + "`r`n")
                $textBoxResults.AppendText("Operation cancelled by user at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n")
            }
            elseif ($script:currentJob.State -eq 'Failed') {
                $textBoxResults.AppendText("`r`n" + "=" * 80 + "`r`n")
                $textBoxResults.AppendText("Operation failed: $($script:currentJob.ChildJobs[0].JobStateInfo.Reason.Message)`r`n")
            }
            
            Remove-Job -Job $script:currentJob -Force
            $script:currentJob = $null
            $timer.Stop()
            Set-ButtonState -Running $false
            $FormMain.Cursor = [System.Windows.Forms.Cursors]::Default
        }
    }
})

# DNS Check button click event
$buttonDNS.Add_Click({
    $selectedTarget = Get-TargetAddress
    if ($null -eq $selectedTarget) {
        return
    }
    
    $script:lastOperation = "DNS"
    
    $textBoxResults.Clear()
    $textBoxResults.Text = "Checking DNS servers for $selectedTarget...`r`n"
    $textBoxResults.Text += "=" * 80 + "`r`n`r`n"
    
    Set-ButtonState -Running $true
    $FormMain.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Start DNS check as a background job
    $script:currentJob = Start-Job -ScriptBlock {
        param($target, $isIP, $AddlDNS)
        
        function Get-ActiveDNSServers {
			# Return a hashtable with activeAdapters and uniqueDNSServers.	
			param ()
			
			$results = @{
				"activeAdapters" = @()
				"uniqueDNSServers" = @()
			}
			
            $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
            
            foreach ($adapter in $adapters) {
                $dnsServers = Get-DnsClientServerAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4 | 
                              Where-Object { $_.ServerAddresses.Count -gt 0 }
                
                if ($dnsServers -and $dnsServers.ServerAddresses) {
					$dnsServers = $dnsServers.ServerAddresses
                    $results.uniqueDNSServers += $dnsServers
                } else {
					$dnsServers = $null
				}
				$results.activeAdapters += [PSCustomObject]@{
					"Name" = $adapter.Name
					"InterfaceDescription" = $adapter.InterfaceDescription
					"dnsServers" = $dnsServers
				}
            }
			
			$results.uniqueDNSServers = $results.uniqueDNSServers | Select -Unique

            return $results
        }
        
        function Test-DNSResolution {
			# Return a string result of whether the resolution failed or not.
			
            param (
                [string]$Domain,
                [string]$DNSServer
            )
            
            try {
                $result = Resolve-DnsName -Name $Domain -Server $DNSServer -ErrorAction Stop
                $ipAddresses = $result | Where-Object { $_.Type -eq "A" -or $_.Type -eq "AAAA" } | Select-Object -ExpandProperty IPAddress
                if ($ipAddresses) {
                    return ($ipAddresses -join ", ")
                }
                else {
                    return "No A/AAAA records found"
                }
            }
            catch {
                return "FAILED: $($_.Exception.Message)"
            }
        }
        
        $dnsInfo = Get-ActiveDNSServers
		
        if (($dnsInfo.activeAdapters | Measure).Count -eq 0) {
            Write-Output "No active network adapters found.`r`n"
        } else {
			Write-Output ("-" * 80 + "`r`n")
			foreach ($adapter in $dnsInfo.activeAdapters) {
					Write-Output "`r`n"
                    Write-Output "Active Adapter: $($adapter.Name)`r`n"
                    Write-Output "Description: $($adapter.InterfaceDescription)`r`n"
					Write-Output "DNS Servers: $($adapter.dnsServers -join ', ')`r`n"
			}
			Write-Output ("-" * 80 + "`r`n")
			
			$dnsServers = $dnsInfo.uniqueDNSServers
			if (($dnsServers | Measure).Count -eq 0) {
				Write-Output "WARNING: No DNS servers found on active network adapters.`r`n"
			}
			if (($AddlDNS | Measure).Count -gt 0) {
				$dnsServers = ($dnsServers + $AddlDNS) | Select -Unique
			}
            if ($isIP) {
                Write-Output "NOTE: Input appears to be an IP address. Performing reverse DNS lookup...`r`n`r`n"

				foreach ($dnsServer in $dnsServers) {
					Write-Output "  DNS Server: $dnsServer`r`n"
					
					try {
						$ptrResult = Resolve-DnsName -Name $target -Server $dnsServer -Type PTR -ErrorAction Stop
						$hostname = $ptrResult | Select-Object -First 1 -ExpandProperty NameHost
						Write-Output "    Reverse Lookup: SUCCESS - Hostname: $hostname`r`n"
					}
					catch {
						Write-Output "    Reverse Lookup: FAILED - $($_.Exception.Message)`r`n"
					}
				}
            } else {	
				foreach ($dnsServer in $dnsServers) {
					Write-Output "  DNS Server: $dnsServer`r`n"
					
					$resolution = Test-DNSResolution -Domain $target -DNSServer $dnsServer
					
					if ($resolution -match "FAILED") {
						Write-Output "    Resolution: $resolution`r`n"
					}
					else {
						Write-Output "    Resolution: SUCCESS - IP(s): $resolution`r`n"
					}
				}
            }
        }
    } -ArgumentList $selectedTarget, (Test-IsIPAddress -Address $selectedTarget), $script:AddlPublicDNS
    
    $timer.Start()
})

# Traceroute button click event
$buttonTraceroute.Add_Click({
    $selectedTarget = Get-TargetAddress
    if ($null -eq $selectedTarget) {
        return
    }
    
    $script:lastOperation = "Tracert"
    
    $textBoxResults.Clear()
    $textBoxResults.Text = "Performing traceroute to $selectedTarget...`r`n"
    $textBoxResults.Text += "=" * 80 + "`r`n`r`n"
    
    Set-ButtonState -Running $true
    $FormMain.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Start traceroute as a background job
    $script:currentJob = Start-Job -ScriptBlock {
        param($target)
        
        Write-Output "Tracing route to $target...`r`n`r`n"
        
        # Create process start info
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = "tracert.exe"
        $psi.Arguments = $target
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true
        
        # Start the process
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $process.Start() | Out-Null
        
        # Read output line by line as it becomes available
        while (-not $process.StandardOutput.EndOfStream) {
            $line = $process.StandardOutput.ReadLine()
            Write-Output "$line`r`n"
        }
        
        # Wait for process to complete
        $process.WaitForExit()
        
        # Check for any errors
        if ($process.ExitCode -ne 0) {
            $errorOutput = $process.StandardError.ReadToEnd()
            if ($errorOutput) {
                Write-Output "`r`nError: $errorOutput`r`n"
            }
        }
        
        $process.Dispose()
        
    } -ArgumentList $selectedTarget
    
    $timer.Start()
})

# Flush DNS button click event
$buttonFlushDNS.Add_Click({
    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will flush the DNS resolver cache. Continue?",
        "Flush DNS Cache",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $textBoxResults.Clear()
        $textBoxResults.Text = "Flushing DNS cache...`r`n"
        $textBoxResults.Text += "=" * 80 + "`r`n`r`n"
        
        try {
            # Run ipconfig /flushdns
            $output = & ipconfig /flushdns 2>&1
            $textBoxResults.AppendText($output -join "`r`n")
            $textBoxResults.AppendText("`r`n`r`n")
            $textBoxResults.AppendText("=" * 80 + "`r`n")
            $textBoxResults.AppendText("DNS cache flushed successfully at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')`r`n")
            
            [System.Windows.Forms.MessageBox]::Show(
                "DNS cache has been flushed successfully.",
                "Success",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            $textBoxResults.AppendText("ERROR: $($_.Exception.Message)`r`n")
            
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to flush DNS cache. Error: $($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

# Adapters button click event
$buttonAdapterInfo.Add_Click({
	$script:lastOperation = "Adapters"
    
	$FormMain.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
	Set-ButtonState -Running $true
	
    $textBoxResults.Clear()
    $textBoxResults.Text = "Getting network adapter info...`r`n"
	
	$script:currentJob = Start-Job -ScriptBlock {
		$adapters = Get-NetAdapter | Select *,@{N="IsUp"; Expression = { if ($_.Status -eq 'Up') { 1 } else { 0 } } } | Sort -Property 'IsUp' -Descending
		foreach ($adapter in $adapters) {
			if ($adapter.InterfaceIndex -ne $null) {
				$config = Get-NetIPConfiguration -InterfaceIndex $adapter.InterfaceIndex -Detailed
			}
			Write-Output "`r`n"
			Write-Output ("=" * 60 + "`r`n")
			Write-Output "Adapter: $($adapter.Name)`r`n"
			Write-Output "NetProfile.Name: $($config.NetProfile.Name)`r`n"
			Write-Output "Description: $($adapter.InterfaceDescription)`r`n"
			Write-Output "Status: $($adapter.Status)`r`n"
			Write-Output ("=" * 60 + "`r`n")
			Write-Output "MAC Address: $($adapter.MacAddress)`r`n"
			Write-Output "Link Speed: $($adapter.LinkSpeed)`r`n"
			Write-Output "Virtual: $($adapter.Virtual)`r`n"
			Write-Output "InterfaceIndex: $($adapter.InterfaceIndex)`r`n"
			Write-Output "Driver Information: $($adapter.DriverInformation)`r`n"
			Write-Output "Driver FileName: $($adapter.DriverFileName)`r`n"
			
			if ($config) {
				Write-Output "`r`n"
				Write-Output "IPv4 Address: $($config.IPv4Address.IPAddress)`r`n"
				Write-Output "IPv4 Default Gateway: $($config.IPv4DefaultGateway.NextHop)`r`n"
				$dnsServers = ($config.DNSServer | where {$_.AddressFamily -eq 2} | Select -ExpandProperty ServerAddresses) -join ', '
				Write-Output "IPv4 DNS Servers: $($dnsServers)`r`n"
				Write-Output "NetProfile.IPv4Connectivity: $($config.NetProfile.IPv4Connectivity)`r`n"
				Write-Output "NetIPv4Interface.DHCP: $($config.NetIPv4Interface.DHCP)`r`n"
				Write-Output "IPv4 Default Gateway Destination Prefix: $($config.IPv4DefaultGateway.DestinationPrefix)`r`n"
				Write-Output "IPv4 Default Gateway Route Metric: $($config.IPv4DefaultGateway.RouteMetric)`r`n"
				Write-Output "IPv4 Default Gateway Interface Metric: $($config.IPv4DefaultGateway.InterfaceMetric)`r`n"
				Write-Output "`r`n"
				Write-Output "IPv6 Address: $($config.IPv6Address.IPAddress)`r`n"
				Write-Output "IPv6 Default Gateway: $($config.IPv6DefaultGateway.NextHop)`r`n"
				$dnsServers = ($config.DNSServer | where {$_.AddressFamily -eq 23} | Select -ExpandProperty ServerAddresses) -join ', '
				Write-Output "IPv6 DNS Servers: $($dnsServers)`r`n"
				Write-Output "NetProfile.IPv6Connectivity: $($config.NetProfile.IPv6Connectivity)`r`n"
				Write-Output "NetIPv6Interface.DHCP: $($config.NetIPv6Interface.DHCP)`r`n"
				Write-Output "IPv6 Default Gateway Destination Prefix: $($config.IPv6DefaultGateway.DestinationPrefix)`r`n"
				Write-Output "IPv6 Default Gateway RouteMetric: $($config.IPv6DefaultGateway.RouteMetric)`r`n"
				Write-Output "IPv6 Default Gateway InterfaceMetric: $($config.IPv6DefaultGateway.InterfaceMetric)`r`n"
			}
		}
	}
	
	$timer.Start()
})

# Save Log button click event
$buttonSaveLog.Add_Click({
    if ([string]::IsNullOrWhiteSpace($textBoxResults.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "No log data to save. Please run a diagnostic operation first.",
            "No Data",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    # Create SaveFileDialog
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    
    # Generate default filename based on last operation
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $defaultFileName = "NetDiagnostics_$($script:lastOperation)_$timestamp.txt"
    
    # Set dialog properties
    $saveFileDialog.FileName = $defaultFileName
    $saveFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save Network Diagnostics Log"
    $saveFileDialog.DefaultExt = "txt"
    
    # Show dialog and save if user clicks OK
    $result = $saveFileDialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            # Save the content to file
            $textBoxResults.Text | Out-File -FilePath $saveFileDialog.FileName -Encoding UTF8
            
            [System.Windows.Forms.MessageBox]::Show(
                "Log saved successfully to:`r`n$($saveFileDialog.FileName)",
                "Save Successful",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to save log file. Error: $($_.Exception.Message)",
                "Save Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
    }
})

# Cancel button click event
$buttonCancel.Add_Click({
    if ($script:currentJob -ne $null) {
        Stop-Job -Job $script:currentJob
        $textBoxResults.AppendText("`r`n[Cancelling operation...]`r`n")
    }
})

# Clear button click event
$buttonClear.Add_Click({
    $textBoxResults.Clear()
})

# Handle Ctrl+C in the form
$FormMain.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq 'C') {
        if ($script:currentJob -ne $null) {
            Stop-Job -Job $script:currentJob
            $textBoxResults.AppendText("`r`n[Operation cancelled with Ctrl+C]`r`n")
        }
    }
})

# Cleanup on form close
$FormMain.Add_FormClosing({
    if ($script:currentJob -ne $null) {
        Stop-Job -Job $script:currentJob
        Remove-Job -Job $script:currentJob -Force
    }
    $timer.Stop()
    $timer.Dispose()
	
	try {
		Stop-Transcript | Out-Null
	} catch {}
})

# Show the form
[void]$FormMain.ShowDialog()

try {
	Stop-Transcript | Out-Null
} catch {}