# Interactive network diagnostic tool
# - nslookup
# - 

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define Hopkins domains
$hopkinsDomains = @(
	"vpn.jh.edu",
    "jhu.edu",
    "johnshopkins.edu",
    "hopkinsmedicine.org",
	"login.johnshopkins.edu",
	"login.microsoftonline.com",
	"microsoft.com"
)

# Define additional name servers to check against.
$addlNameservers = @(
	'8.8.8.8',
	'1.1.1.1'
)

# Global variable to track running job and last operation
$script:currentJob = $null
$script:lastOperation = "DNS"  # Default operation type for filename

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "IT Network Diagnostics"
$form.Size = New-Object System.Drawing.Size(800, 600)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false

# Define margins and spacing
$marginLeft = 10
$marginTop = 10
$spacing = 10
$buttonHeight = 30
$labelHeight = 20
$currentY = $marginTop

# Create domain selection label
$labelDomain = New-Object System.Windows.Forms.Label
$labelDomain.Location = New-Object System.Drawing.Point($marginLeft, $currentY)
$labelDomain.Size = New-Object System.Drawing.Size(150, $labelHeight)
$labelDomain.Text = "Select Hopkins Domain:"
$form.Controls.Add($labelDomain)

# Create domain dropdown (positioned to the right of label)
$comboBoxDomain = New-Object System.Windows.Forms.ComboBox
$comboBoxDomain.Location = New-Object System.Drawing.Point(($labelDomain.Right + $spacing), $currentY)
$comboBoxDomain.Size = New-Object System.Drawing.Size(200, $labelHeight)
$comboBoxDomain.DropDownStyle = "DropDownList"
foreach ($domain in $hopkinsDomains) {
    $comboBoxDomain.Items.Add($domain) | Out-Null
}
$comboBoxDomain.Items.Add("-- Custom Address --") | Out-Null
$comboBoxDomain.SelectedIndex = 0
$form.Controls.Add($comboBoxDomain)

# Create custom address label (positioned to the right of dropdown)
$labelCustom = New-Object System.Windows.Forms.Label
$labelCustom.Location = New-Object System.Drawing.Point(($comboBoxDomain.Right + 20), $currentY)
$labelCustom.Size = New-Object System.Drawing.Size(100, $labelHeight)
$labelCustom.Text = "Custom Address:"
$labelCustom.Visible = $false
$form.Controls.Add($labelCustom)

# Create custom address textbox (positioned to the right of custom label)
$textBoxCustom = New-Object System.Windows.Forms.TextBox
$textBoxCustom.Location = New-Object System.Drawing.Point(($labelCustom.Right + $spacing), $currentY)
$textBoxCustom.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - $labelCustom.Right - $spacing - $marginLeft), $labelHeight)
$textBoxCustom.Visible = $false
$textBoxCustom.Text = ""
$form.Controls.Add($textBoxCustom)

# Move to next row
$currentY = $labelDomain.Bottom + $spacing

# First row of buttons
# Create DNS Check button
$buttonDNS = New-Object System.Windows.Forms.Button
$buttonDNS.Location = New-Object System.Drawing.Point($marginLeft, $currentY)
$buttonDNS.Size = New-Object System.Drawing.Size(140, $buttonHeight)
$buttonDNS.Text = "Check DNS Servers"
$buttonDNS.BackColor = [System.Drawing.Color]::LightBlue
$form.Controls.Add($buttonDNS)

# Create Traceroute button (positioned to the right of DNS button)
$buttonTraceroute = New-Object System.Windows.Forms.Button
$buttonTraceroute.Location = New-Object System.Drawing.Point(($buttonDNS.Right + $spacing), $currentY)
$buttonTraceroute.Size = New-Object System.Drawing.Size(140, $buttonHeight)
$buttonTraceroute.Text = "Perform Traceroute"
$buttonTraceroute.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($buttonTraceroute)

# Create Flush DNS button (positioned to the right of Traceroute button)
$buttonFlushDNS = New-Object System.Windows.Forms.Button
$buttonFlushDNS.Location = New-Object System.Drawing.Point(($buttonTraceroute.Right + $spacing), $currentY)
$buttonFlushDNS.Size = New-Object System.Drawing.Size(100, $buttonHeight)
$buttonFlushDNS.Text = "Flush DNS"
$buttonFlushDNS.BackColor = [System.Drawing.Color]::LightYellow
$form.Controls.Add($buttonFlushDNS)

# Create Cancel button (positioned to the right of Flush DNS button)
$buttonCancel = New-Object System.Windows.Forms.Button
$buttonCancel.Location = New-Object System.Drawing.Point(($buttonFlushDNS.Right + $spacing), $currentY)
$buttonCancel.Size = New-Object System.Drawing.Size(100, $buttonHeight)
$buttonCancel.Text = "Cancel"
$buttonCancel.BackColor = [System.Drawing.Color]::Orange
$buttonCancel.Enabled = $false
$form.Controls.Add($buttonCancel)

# Create Save Log button (positioned to the right of Cancel button)
$buttonSaveLog = New-Object System.Windows.Forms.Button
$buttonSaveLog.Location = New-Object System.Drawing.Point(($buttonCancel.Right + $spacing), $currentY)
$buttonSaveLog.Size = New-Object System.Drawing.Size(100, $buttonHeight)
$buttonSaveLog.Text = "Save Log"
$buttonSaveLog.BackColor = [System.Drawing.Color]::LightCyan
$form.Controls.Add($buttonSaveLog)

# Create Clear button (positioned to the right of Save Log button)
$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Location = New-Object System.Drawing.Point(($buttonSaveLog.Right + $spacing), $currentY)
$buttonClear.Size = New-Object System.Drawing.Size(100, $buttonHeight)
$buttonClear.Text = "Clear Results"
$buttonClear.BackColor = [System.Drawing.Color]::LightCoral
$form.Controls.Add($buttonClear)

# Move to next row for results textbox
$currentY = $buttonDNS.Bottom + $spacing

# Create results text box (fills remaining space)
$textBoxResults = New-Object System.Windows.Forms.TextBox
$textBoxResults.Location = New-Object System.Drawing.Point($marginLeft, $currentY)
$textBoxResults.Size = New-Object System.Drawing.Size(
    ($form.ClientSize.Width - (2 * $marginLeft)),
    ($form.ClientSize.Height - $currentY - $marginLeft)
)
$textBoxResults.Multiline = $true
$textBoxResults.ScrollBars = "Vertical"
$textBoxResults.Font = New-Object System.Drawing.Font("Consolas", 9)
$textBoxResults.ReadOnly = $true
$textBoxResults.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor 
                         [System.Windows.Forms.AnchorStyles]::Bottom -bor 
                         [System.Windows.Forms.AnchorStyles]::Left -bor 
                         [System.Windows.Forms.AnchorStyles]::Right
$form.Controls.Add($textBoxResults)

# Create a timer for updating job output
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 250  # Update every 250ms

# Event handler for dropdown selection change
$comboBoxDomain.Add_SelectedIndexChanged({
    if ($comboBoxDomain.SelectedItem -eq "-- Custom Address --") {
        $labelCustom.Visible = $true
        $textBoxCustom.Visible = $true
        $textBoxCustom.Focus()
    }
    else {
        $labelCustom.Visible = $false
        $textBoxCustom.Visible = $false
    }
})

# Function to get the target address (either from dropdown or custom input)
function Get-TargetAddress {
    if ($comboBoxDomain.SelectedItem -eq "-- Custom Address --") {
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
        return $comboBoxDomain.SelectedItem
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
    $comboBoxDomain.Enabled = -not $Running
    $textBoxCustom.Enabled = -not $Running
}

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
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
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
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
    # Start DNS check as a background job
    $script:currentJob = Start-Job -ScriptBlock {
        param($target, $isIP)
        
        function Get-ActiveDNSServers {
            $results = @()
            $adapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
            
            foreach ($adapter in $adapters) {
                $dnsServers = Get-DnsClientServerAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4 | 
                              Where-Object { $_.ServerAddresses.Count -gt 0 }
                
                if ($dnsServers -and $dnsServers.ServerAddresses) {
                    $results += [PSCustomObject]@{
                        AdapterName = $adapter.Name
                        InterfaceDescription = $adapter.InterfaceDescription
                        DNSServers = $dnsServers.ServerAddresses
                    }
                }
            }
            
            return $results
        }
        
        function Test-DNSResolution {
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
        
        if ($dnsInfo.Count -eq 0) {
            Write-Output "No active network adapters with DNS servers found.`r`n"
        }
        else {
			if (($addlNameServers | Measure).Count -gt 0) {
				$dnsInfo += [PSCustomObject]@{
						AdapterName = "N/A"
						InterfaceDescription = "Not an adapter - testing against public DNS servers"
						DNSServers = $addlNameServers
				}
			}
				
            if ($isIP) {
                Write-Output "NOTE: Input appears to be an IP address. Performing reverse DNS lookup...`r`n`r`n"
                
                foreach ($adapter in $dnsInfo) {
                    Write-Output "Adapter: $($adapter.AdapterName)`r`n"
                    Write-Output "Description: $($adapter.InterfaceDescription)`r`n"
                    Write-Output ("-" * 80 + "`r`n")
                    
                    foreach ($dnsServer in $adapter.DNSServers) {
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
                    
                    Write-Output "`r`n"
                }
            }
            else {	
                foreach ($adapter in $dnsInfo) {
                    Write-Output "Adapter: $($adapter.AdapterName)`r`n"
                    Write-Output "Description: $($adapter.InterfaceDescription)`r`n"
                    Write-Output ("-" * 80 + "`r`n")
                    
                    foreach ($dnsServer in $adapter.DNSServers) {
                        Write-Output "  DNS Server: $dnsServer`r`n"
                        
                        $resolution = Test-DNSResolution -Domain $target -DNSServer $dnsServer
                        
                        if ($resolution -match "FAILED") {
                            Write-Output "    Resolution: $resolution`r`n"
                        }
                        else {
                            Write-Output "    Resolution: SUCCESS - IP(s): $resolution`r`n"
                        }
                    }
                    
                    Write-Output "`r`n"
                }
            }
        }
        
    } -ArgumentList $selectedTarget, (Test-IsIPAddress -Address $selectedTarget)
    
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
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    
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
$form.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq 'C') {
        if ($script:currentJob -ne $null) {
            Stop-Job -Job $script:currentJob
            $textBoxResults.AppendText("`r`n[Operation cancelled with Ctrl+C]`r`n")
        }
    }
})

# Cleanup on form close
$form.Add_FormClosing({
    if ($script:currentJob -ne $null) {
        Stop-Job -Job $script:currentJob
        Remove-Job -Job $script:currentJob -Force
    }
    $timer.Stop()
    $timer.Dispose()
})

# Show the form
[void]$form.ShowDialog()