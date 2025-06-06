Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# First, include the Get-LastInstalledUpdate function
function Get-LastInstalledUpdate {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true, 
                   ValueFromPipelineByPropertyName = $true)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        
        [Parameter()]
        [System.Management.Automation.PSCredential]
        [System.Management.Automation.Credential()]
        $Credential = [System.Management.Automation.PSCredential]::Empty,
        
        [Parameter()]
        [switch]$UseSSL
    )
    
    begin {
        # Set up the session options to bypass proxy
        $SessionOption = New-PSSessionOption -ProxyAccessType NoProxyServer
    }
    
    process {
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Checking last installed update on $Computer"
            
            $InvokeParams = @{
                ComputerName = $Computer
                ScriptBlock  = {
                    # Get all installed Windows updates to allow filtering by type
                    Get-HotFix | Sort-Object -Property InstalledOn -Descending
                }
                ErrorAction  = 'Stop'
                SessionOption = $SessionOption
            }
            
            # Add UseSSL if specified
            if ($UseSSL) {
                $InvokeParams.Add('UseSSL', $true)
            }
            
            # Add credentials if provided
            if ($Credential -ne [System.Management.Automation.PSCredential]::Empty) {
                $InvokeParams.Add('Credential', $Credential)
            }
            
            try {
                $Result = Invoke-Command @InvokeParams
                
                # Return results (can be multiple hotfixes)
                foreach ($hotfix in $Result) {
                    [PSCustomObject]@{
                        ComputerName = $Computer
                        HotFixID     = $hotfix.HotFixID
                        Description  = $hotfix.Description
                        InstalledOn  = $hotfix.InstalledOn
                        InstalledBy  = $hotfix.InstalledBy
                        Status       = "Success"
                    }
                }
            }
            catch {
                # Return error information in the same object format
                [PSCustomObject]@{
                    ComputerName = $Computer
                    HotFixID     = $null
                    Description  = $null
                    InstalledOn  = $null
                    InstalledBy  = $null
                    Status       = "Error: $_"
                }
            }
        }
    }
}

# Variables to store data
$script:importedServers = @()
$script:results = @()
$script:filterButtonsCreated = $false

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Server Update Checker"
$form.Size = New-Object System.Drawing.Size(1300, 650)  # Increased height for better visibility
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)

# Create the Import CSV button
$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Location = New-Object System.Drawing.Point(20, 20)
$btnImport.Size = New-Object System.Drawing.Size(150, 30)
$btnImport.Text = "Import CSV"
$form.Controls.Add($btnImport)

# Create label to show CSV path
$lblCSVPath = New-Object System.Windows.Forms.Label
$lblCSVPath.Location = New-Object System.Drawing.Point(180, 25)
$lblCSVPath.Size = New-Object System.Drawing.Size(600, 20)
$lblCSVPath.Text = "No CSV file selected"
$form.Controls.Add($lblCSVPath)

# Create servers list box
$lblServers = New-Object System.Windows.Forms.Label
$lblServers.Location = New-Object System.Drawing.Point(20, 60)
$lblServers.Size = New-Object System.Drawing.Size(150, 20)
$lblServers.Text = "Imported Servers:"
$form.Controls.Add($lblServers)

$listServers = New-Object System.Windows.Forms.ListBox
$listServers.Location = New-Object System.Drawing.Point(20, 85)
$listServers.Size = New-Object System.Drawing.Size(250, 430)
$listServers.SelectionMode = "MultiExtended"
$form.Controls.Add($listServers)

# Create Options GroupBox
$gbOptions = New-Object System.Windows.Forms.GroupBox
$gbOptions.Location = New-Object System.Drawing.Point(290, 85)
$gbOptions.Size = New-Object System.Drawing.Size(250, 150)
$gbOptions.Text = "Connection Options"
$form.Controls.Add($gbOptions)

# Create SSL Checkbox
$chkSSL = New-Object System.Windows.Forms.CheckBox
$chkSSL.Location = New-Object System.Drawing.Point(20, 30)
$chkSSL.Size = New-Object System.Drawing.Size(200, 20)
$chkSSL.Text = "Use SSL"
$gbOptions.Controls.Add($chkSSL)

# Create Credentials Checkbox
$chkCreds = New-Object System.Windows.Forms.CheckBox
$chkCreds.Location = New-Object System.Drawing.Point(20, 60)
$chkCreds.Size = New-Object System.Drawing.Size(200, 20)
$chkCreds.Text = "Use Custom Credentials"
$gbOptions.Controls.Add($chkCreds)

# Create Filter Options GroupBox
$gbFilter = New-Object System.Windows.Forms.GroupBox
$gbFilter.Location = New-Object System.Drawing.Point(290, 350)
$gbFilter.Size = New-Object System.Drawing.Size(250, 165)  # Increased height for the new checkbox
$gbFilter.Text = "Filter by Update Type"
$form.Controls.Add($gbFilter)

# Common update types
$script:updateTypes = @(
    "Update",
    "Security Update",
    "Hotfix",
    "Service Pack",
    "Critical Update"
)

# Create checkboxes for update types
$checkboxY = 25
$script:typeCheckboxes = @{}

foreach ($type in $script:updateTypes) {
    $checkbox = New-Object System.Windows.Forms.CheckBox
    $checkbox.Location = New-Object System.Drawing.Point(20, $checkboxY)
    $checkbox.Size = New-Object System.Drawing.Size(200, 20)
    $checkbox.Text = $type
    $checkbox.Checked = $true  # Default to checked
    $gbFilter.Controls.Add($checkbox)
    $script:typeCheckboxes[$type] = $checkbox
    $checkboxY += 22
}

# Add "Show Only Latest Update" checkbox
$chkLatestOnly = New-Object System.Windows.Forms.CheckBox
$chkLatestOnly.Location = New-Object System.Drawing.Point(20, ($checkboxY + 5))
$chkLatestOnly.Size = New-Object System.Drawing.Size(220, 20)
$chkLatestOnly.Text = "Latest Update Only"
$chkLatestOnly.Checked = $false
$gbFilter.Controls.Add($chkLatestOnly)

# Create Get Updates button
$btnGetUpdates = New-Object System.Windows.Forms.Button
$btnGetUpdates.Location = New-Object System.Drawing.Point(290, 250)
$btnGetUpdates.Size = New-Object System.Drawing.Size(250, 40)
$btnGetUpdates.Text = "Get Last Updates"
$btnGetUpdates.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 215)
$btnGetUpdates.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnGetUpdates)

# Create Export Results button
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Location = New-Object System.Drawing.Point(290, 300)
$btnExport.Size = New-Object System.Drawing.Size(250, 40)
$btnExport.Text = "Export Results"
$btnExport.Enabled = $false
$form.Controls.Add($btnExport)

# Create "Apply Filters" button
$btnFilter = New-Object System.Windows.Forms.Button
$btnFilter.Location = New-Object System.Drawing.Point(560, 530)  # Moved to match progress bar
$btnFilter.Size = New-Object System.Drawing.Size(200, 30)
$btnFilter.Text = "Apply Filters"
$btnFilter.Enabled = $false
$form.Controls.Add($btnFilter)

# Create "Show All" button
$btnShowAll = New-Object System.Windows.Forms.Button
$btnShowAll.Location = New-Object System.Drawing.Point(770, 530)  # Moved to match progress bar
$btnShowAll.Size = New-Object System.Drawing.Size(200, 30)
$btnShowAll.Text = "Show All Results"
$btnShowAll.Enabled = $false
$form.Controls.Add($btnShowAll)

# Create Results DataGridView
$lblResults = New-Object System.Windows.Forms.Label
$lblResults.Location = New-Object System.Drawing.Point(560, 60)
$lblResults.Size = New-Object System.Drawing.Size(150, 20)
$lblResults.Text = "Results:"
$form.Controls.Add($lblResults)

$dataGrid = New-Object System.Windows.Forms.DataGridView
$dataGrid.Location = New-Object System.Drawing.Point(560, 85)
$dataGrid.Size = New-Object System.Drawing.Size(710, 430)
$dataGrid.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dataGrid.ReadOnly = $true
$dataGrid.AllowUserToAddRows = $false
$dataGrid.AllowUserToDeleteRows = $false
$dataGrid.AllowUserToOrderColumns = $true
$dataGrid.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dataGrid.MultiSelect = $false
$dataGrid.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$form.Controls.Add($dataGrid)

# Status bar
$statusBar = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusBar.Items.Add($statusLabel)
$form.Controls.Add($statusBar)

# Filter button click event
$btnFilter.Add_Click({
    if ($script:results.Count -eq 0) {
        return
    }
    
    # Apply filtering based on selected Description types and latest-only option
    $filteredResults = $script:results | Where-Object {
        if ([string]::IsNullOrEmpty($_.Description)) {
            # Show results with no description - likely errors
            return $true
        }
        
        # Check if this Description type is selected in checkboxes
        $matchFound = $false
        
        foreach ($type in $script:updateTypes) {
            # For "Update", make sure it's not part of another update type
            if ($type -eq "Update" -and $_.Description -eq "Update" -and $script:typeCheckboxes[$type].Checked) {
                $matchFound = $true
                break
            }
            # For other update types, use a more specific match
            elseif ($type -ne "Update" -and $_.Description -eq $type -and $script:typeCheckboxes[$type].Checked) {
                $matchFound = $true
                break
            }
        }
        
        if ($matchFound) {
            return $true
        }
        
        # If none of our known types match, show it by default (unknown update type)
        $matchedKnownType = $false
        foreach ($type in $script:updateTypes) {
            if ($_.Description -eq $type) {
                $matchedKnownType = $true
                break
            }
        }
        
        # Show if it's an unknown type
        return -not $matchedKnownType
    }
    
    # Apply "Show Only Latest Update Per Server" filter if checked
    if ($chkLatestOnly.Checked) {
        $latestByComputer = @{}
        
        # Group by computer and find the latest update for each
        foreach ($result in $filteredResults) {
            $computer = $result.ComputerName
            
            # Skip entries with no installation date (likely errors)
            if ($null -eq $result.InstalledOn) {
                continue
            }
            
            # Initialize or update the latest update for this computer
            if (-not $latestByComputer.ContainsKey($computer) -or 
                $result.InstalledOn -gt $latestByComputer[$computer].InstalledOn) {
                $latestByComputer[$computer] = $result
            }
        }
        
        # Convert the hashtable values to an array
        $filteredResults = $latestByComputer.Values
    }
    
    # Update the DataGridView
    $dataGrid.DataSource = [System.Collections.ArrayList]($filteredResults)
    $statusLabel.Text = "Filtered to show $($filteredResults.Count) of $($script:results.Count) results"
})

# Show All button click event
$btnShowAll.Add_Click({
    if ($script:results.Count -eq 0) {
        return
    }
    
    # Show all results
    $dataGrid.DataSource = [System.Collections.ArrayList]($script:results)
    $statusLabel.Text = "Showing all $($script:results.Count) results"
})

# Import CSV button click event
$btnImport.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select a CSV file with server names"
    
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $csvPath = $openFileDialog.FileName
        $lblCSVPath.Text = $csvPath
        
        try {
            $statusLabel.Text = "Importing servers from CSV..."
            
            # Import servers from CSV - expecting a column named "ComputerName" or "ServerName"
            $csvData = Import-Csv -Path $csvPath
            $script:importedServers = @()
            
            # Try to find the server name column
            $serverNameColumn = $null
            if ($csvData[0].PSObject.Properties.Name -contains "ComputerName") {
                $serverNameColumn = "ComputerName"
            }
            elseif ($csvData[0].PSObject.Properties.Name -contains "ServerName") {
                $serverNameColumn = "ServerName"
            }
            elseif ($csvData[0].PSObject.Properties.Name -contains "Server") {
                $serverNameColumn = "Server"
            }
            elseif ($csvData[0].PSObject.Properties.Name -contains "Name") {
                $serverNameColumn = "Name"
            }
            elseif ($csvData[0].PSObject.Properties.Name -contains "Computer") {
                $serverNameColumn = "Computer"
            }
            else {
                # If no obvious column name, use the first column
                $serverNameColumn = $csvData[0].PSObject.Properties.Name[0]
            }
            
            foreach ($row in $csvData) {
                $serverName = $row.$serverNameColumn
                if (-not [string]::IsNullOrWhiteSpace($serverName)) {
                    $script:importedServers += $serverName
                }
            }
            
            # Update the list box
            $listServers.Items.Clear()
            foreach ($server in $script:importedServers) {
                $listServers.Items.Add($server)
            }
            
            $statusLabel.Text = "Imported $($script:importedServers.Count) servers from CSV"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error importing CSV: $_", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error importing CSV"
        }
    }
})

# Get Updates button click event
$btnGetUpdates.Add_Click({
    # Make sure we have servers selected
    if ($listServers.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one server from the list", "No Servers Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $selectedServers = $listServers.SelectedItems
    $statusLabel.Text = "Querying updates for $($selectedServers.Count) servers..."
    
    # Create a progress bar
    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(290, 530)  # Moved lower for better visibility
    $progressBar.Size = New-Object System.Drawing.Size(250, 20)
    $progressBar.Minimum = 0
    $progressBar.Maximum = $selectedServers.Count
    $progressBar.Value = 0
    $form.Controls.Add($progressBar)
    $form.Update()
    
    # Parameters for Get-LastInstalledUpdate
    $params = @{}
    
    # Check if using SSL
    if ($chkSSL.Checked) {
        $params.Add('UseSSL', $true)
    }
    
    # Check if using custom credentials
    if ($chkCreds.Checked) {
        $cred = Get-Credential -Message "Enter credentials for remote server access"
        if ($cred) {
            $params.Add('Credential', $cred)
        }
        else {
            $progressBar.Dispose()
            $statusLabel.Text = "Credentials not provided, operation cancelled"
            return
        }
    }
    
    # Execute the function
    $script:results = @()
    
    foreach ($server in $selectedServers) {
        try {
            $statusLabel.Text = "Querying $server..."
            $results = Get-LastInstalledUpdate -ComputerName $server @params
            $script:results += $results
            
            # Update progress
            $progressBar.Value++
            $form.Update()
        }
        catch {
            $script:results += [PSCustomObject]@{
                ComputerName = $server
                HotFixID     = $null
                Description  = $null
                InstalledOn  = $null
                InstalledBy  = $null
                Status       = "Error: $_"
            }
        }
    }
    
    # Update the DataGridView with all results initially
    $dataGrid.DataSource = [System.Collections.ArrayList]($script:results)
    
    # Clean up
    $progressBar.Dispose()
    $statusLabel.Text = "Completed querying $($selectedServers.Count) servers"
    
    # Enable buttons if we have results
    if ($script:results.Count -gt 0) {
        $btnExport.Enabled = $true
        $btnFilter.Enabled = $true
        $btnShowAll.Enabled = $true
    }
})

# Export Results button click event
$btnExport.Add_Click({
    if ($script:results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No results to export", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $saveFileDialog.Title = "Save results as CSV"
    $saveFileDialog.FileName = "UpdateResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $script:results | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
            $statusLabel.Text = "Results exported to $($saveFileDialog.FileName)"
            [System.Windows.Forms.MessageBox]::Show("Results exported successfully!", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Error exporting results: $_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            $statusLabel.Text = "Error exporting results"
        }
    }
})

# Show the form
[void]$form.ShowDialog()