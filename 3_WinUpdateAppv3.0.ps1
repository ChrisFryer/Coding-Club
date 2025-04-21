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

# Create "Produce Report" button
$btnReport = New-Object System.Windows.Forms.Button
$btnReport.Location = New-Object System.Drawing.Point(980, 530)  # Next to Show All Results button
$btnReport.Size = New-Object System.Drawing.Size(200, 30)
$btnReport.Text = "Produce Report"
$btnReport.Enabled = $false
$form.Controls.Add($btnReport)

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

# Function to create a report
function Show-UpdateReport {
    param(
        [array]$Results
    )

    # Create a new form for the report
    $reportForm = New-Object System.Windows.Forms.Form
    $reportForm.Text = "Patch Status Report"
    $reportForm.Size = New-Object System.Drawing.Size(1000, 700)
    $reportForm.StartPosition = "CenterScreen"
    $reportForm.Font = New-Object System.Drawing.Font("Segoe UI", 10)

    # Add title
    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Location = New-Object System.Drawing.Point(20, 20)
    $lblTitle.Size = New-Object System.Drawing.Size(960, 30)
    $lblTitle.Text = "Server Patch Status Report - Generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $reportForm.Controls.Add($lblTitle)

    # Get the current and previous month
    $currentMonth = (Get-Date).Month
    $currentYear = (Get-Date).Year
    $previousMonth = (Get-Date).AddMonths(-1).Month
    $previousYear = (Get-Date).AddMonths(-1).Year

    # Analyze the data for reporting
    $serverStatus = @()
    $serverGroups = @{
        "CurrentMonth" = @() # Patched in current month (up-to-date)
        "PreviousMonth" = @() # Patched in previous month
        "Older" = @() # Patched before previous month
        "NoPatches" = @() # No patches or errors
    }

    $monthlyPatchCounts = @{}

    # Process each server
    $uniqueServers = $Results | Select-Object -ExpandProperty ComputerName -Unique

    foreach ($server in $uniqueServers) {
        $serverPatches = $Results | Where-Object { $_.ComputerName -eq $server } | Sort-Object -Property InstalledOn -Descending
        
        if (-not $serverPatches -or $serverPatches.Count -eq 0 -or $null -eq $serverPatches[0].InstalledOn) {
            $serverGroups["NoPatches"] += $server
            $serverStatus += [PSCustomObject]@{
                ServerName = $server
                LastPatchDate = "N/A"
                DaysSinceLastPatch = "N/A"
                Status = "No Data"
            }
            continue
        }
        
        $latestPatch = $serverPatches | Where-Object { $null -ne $_.InstalledOn } | Select-Object -First 1
        $patchDate = $latestPatch.InstalledOn
        $daysSinceLastPatch = (New-TimeSpan -Start $patchDate -End (Get-Date)).Days
        
        # Determine status based on patch date
        if ($patchDate.Month -eq $currentMonth -and $patchDate.Year -eq $currentYear) {
            $status = "Up-to-date"
            $serverGroups["CurrentMonth"] += $server
        }
        elseif ($patchDate.Month -eq $previousMonth -and $patchDate.Year -eq $previousYear) {
            $status = "Last Month"
            $serverGroups["PreviousMonth"] += $server
        }
        else {
            $status = "Outdated"
            $serverGroups["Older"] += $server
        }
        
        $serverStatus += [PSCustomObject]@{
            ServerName = $server
            LastPatchDate = $patchDate.ToString("yyyy-MM-dd")
            DaysSinceLastPatch = $daysSinceLastPatch
            Status = $status
        }
        
        # Count patches by month
        foreach ($patch in $serverPatches) {
            if ($null -eq $patch.InstalledOn) { continue }
            
            $monthYear = "$($patch.InstalledOn.Year)-$($patch.InstalledOn.Month.ToString('00'))"
            if (-not $monthlyPatchCounts.ContainsKey($monthYear)) {
                $monthlyPatchCounts[$monthYear] = 0
            }
            $monthlyPatchCounts[$monthYear]++
        }
    }
    
    # Sort the monthly patch counts
    $sortedMonthlyPatches = $monthlyPatchCounts.GetEnumerator() | Sort-Object -Property Name -Descending | Select-Object -First 6

    # Create a panel for the status summary
    $pnlSummary = New-Object System.Windows.Forms.Panel
    $pnlSummary.Location = New-Object System.Drawing.Point(20, 70)
    $pnlSummary.Size = New-Object System.Drawing.Size(460, 250)
    $pnlSummary.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $reportForm.Controls.Add($pnlSummary)

    # Add summary title
    $lblSummaryTitle = New-Object System.Windows.Forms.Label
    $lblSummaryTitle.Location = New-Object System.Drawing.Point(10, 10)
    $lblSummaryTitle.Size = New-Object System.Drawing.Size(440, 20)
    $lblSummaryTitle.Text = "Patch Status Summary"
    $lblSummaryTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $pnlSummary.Controls.Add($lblSummaryTitle)

    # Add summary metrics
    $summaryY = 40
    $summaryMetrics = @(
        @{Label = "Total Servers Scanned:"; Value = $uniqueServers.Count},
        @{Label = "Up-to-date Servers (Current Month):"; Value = $serverGroups["CurrentMonth"].Count},
        @{Label = "Last Month Patched Servers:"; Value = $serverGroups["PreviousMonth"].Count},
        @{Label = "Outdated Servers:"; Value = $serverGroups["Older"].Count},
        @{Label = "Servers with No Patch Data:"; Value = $serverGroups["NoPatches"].Count}
    )

    foreach ($metric in $summaryMetrics) {
        $lblMetric = New-Object System.Windows.Forms.Label
        $lblMetric.Location = New-Object System.Drawing.Point(20, $summaryY)
        $lblMetric.Size = New-Object System.Drawing.Size(250, 20)
        $lblMetric.Text = $metric.Label
        $pnlSummary.Controls.Add($lblMetric)
        
        $lblValue = New-Object System.Windows.Forms.Label
        $lblValue.Location = New-Object System.Drawing.Point(270, $summaryY)
        $lblValue.Size = New-Object System.Drawing.Size(100, 20)
        $lblValue.Text = $metric.Value
        $lblValue.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
        $pnlSummary.Controls.Add($lblValue)
        
        $summaryY += 30
    }

    # Create a panel for the status distribution chart (pie chart)
    $pnlStatusChart = New-Object System.Windows.Forms.Panel
    $pnlStatusChart.Location = New-Object System.Drawing.Point(500, 70)
    $pnlStatusChart.Size = New-Object System.Drawing.Size(460, 250)
    $pnlStatusChart.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $reportForm.Controls.Add($pnlStatusChart)

    # Add chart title
    $lblChartTitle = New-Object System.Windows.Forms.Label
    $lblChartTitle.Location = New-Object System.Drawing.Point(10, 10)
    $lblChartTitle.Size = New-Object System.Drawing.Size(440, 20)
    $lblChartTitle.Text = "Patch Status Distribution"
    $lblChartTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $pnlStatusChart.Controls.Add($lblChartTitle)

	# Instead of using the Paint property, we need to use the Add_Paint event handler
	# Replace these sections in your script:

	# For the pie chart panel:
	$chartPanel = New-Object System.Windows.Forms.Panel
	$chartPanel.Location = New-Object System.Drawing.Point(10, 40)
	$chartPanel.Size = New-Object System.Drawing.Size(200, 200)
	$pnlStatusChart.Controls.Add($chartPanel)

	# Use Add_Paint instead of Paint property
	$chartPanel.Add_Paint({
		$graphics = $_.Graphics
		$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
		
		$total = $uniqueServers.Count
		if ($total -eq 0) { $total = 1 } # Avoid division by zero
		
		$center = New-Object System.Drawing.Point(100, 100)
		$radius = 90
		$startAngle = 0
		
		# Define colors
		$colors = @(
			[System.Drawing.Color]::FromArgb(92, 184, 92),    # Green - Up-to-date
			[System.Drawing.Color]::FromArgb(240, 173, 78),   # Yellow - Last Month
			[System.Drawing.Color]::FromArgb(217, 83, 79),    # Red - Outdated
			[System.Drawing.Color]::FromArgb(91, 192, 222)    # Blue - No Data
		)
		
		$slices = @(
			@{Count = $serverGroups["CurrentMonth"].Count; Percent = $serverGroups["CurrentMonth"].Count / $total},
			@{Count = $serverGroups["PreviousMonth"].Count; Percent = $serverGroups["PreviousMonth"].Count / $total},
			@{Count = $serverGroups["Older"].Count; Percent = $serverGroups["Older"].Count / $total},
			@{Count = $serverGroups["NoPatches"].Count; Percent = $serverGroups["NoPatches"].Count / $total}
		)
		
		for ($i = 0; $i -lt $slices.Count; $i++) {
			$sweepAngle = 360 * $slices[$i].Percent
			if ($sweepAngle -gt 0) {
				$brush = New-Object System.Drawing.SolidBrush($colors[$i])
				$graphics.FillPie($brush, 10, 10, 180, 180, $startAngle, $sweepAngle)
				$startAngle += $sweepAngle
				$brush.Dispose()
			}
		}
	})

    # Add legend
    $legendY = 40
    $legendX = 220
    $legendItems = @(
        @{Color = [System.Drawing.Color]::FromArgb(92, 184, 92); Label = "Up-to-date (Current Month)"; Count = $serverGroups["CurrentMonth"].Count},
        @{Color = [System.Drawing.Color]::FromArgb(240, 173, 78); Label = "Last Month"; Count = $serverGroups["PreviousMonth"].Count},
        @{Color = [System.Drawing.Color]::FromArgb(217, 83, 79); Label = "Outdated"; Count = $serverGroups["Older"].Count},
        @{Color = [System.Drawing.Color]::FromArgb(91, 192, 222); Label = "No Patch Data"; Count = $serverGroups["NoPatches"].Count}
    )

	# For the legend items fix:
	# Fix the legendX + 30 issue by using explicit integers
	foreach ($item in $legendItems) {
		$pnlColor = New-Object System.Windows.Forms.Panel
		$pnlColor.Location = New-Object System.Drawing.Point($legendX, $legendY)
		$pnlColor.Size = New-Object System.Drawing.Size(20, 20)
		$pnlColor.BackColor = $item.Color
		$pnlStatusChart.Controls.Add($pnlColor)
    
		# Calculate positions explicitly
		$labelX = $legendX + 30 # Use explicit addition
    
		$lblLegend = New-Object System.Windows.Forms.Label
		$lblLegend.Location = New-Object System.Drawing.Point($labelX, $legendY)
		$lblLegend.Size = New-Object System.Drawing.Size(200, 20)
		$lblLegend.Text = "$($item.Label): $($item.Count)"
		$pnlStatusChart.Controls.Add($lblLegend)
    
		$legendY += 30 # This is safe because it's a simple variable
	}

    # Create a panel for monthly patch trend chart (bar chart)
    $pnlTrendChart = New-Object System.Windows.Forms.Panel
    $pnlTrendChart.Location = New-Object System.Drawing.Point(20, 340)
    $pnlTrendChart.Size = New-Object System.Drawing.Size(940, 250)
    $pnlTrendChart.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $reportForm.Controls.Add($pnlTrendChart)

    # Add chart title
    $lblTrendTitle = New-Object System.Windows.Forms.Label
    $lblTrendTitle.Location = New-Object System.Drawing.Point(10, 10)
    $lblTrendTitle.Size = New-Object System.Drawing.Size(920, 20)
    $lblTrendTitle.Text = "Monthly Patch Trend (Last 6 Months)"
    $lblTrendTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $pnlTrendChart.Controls.Add($lblTrendTitle)

    # For the bar chart panel:
	$barChartPanel = New-Object System.Windows.Forms.Panel
	$barChartPanel.Location = New-Object System.Drawing.Point(50, 40)
	$barChartPanel.Size = New-Object System.Drawing.Size(880, 190)
	$pnlTrendChart.Controls.Add($barChartPanel)

	# Use Add_Paint instead of Paint property
	$barChartPanel.Add_Paint({
		$graphics = $_.Graphics
		$graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    
		$chartWidth = $barChartPanel.Width - 100
		$chartHeight = $barChartPanel.Height - 60
    
		# Draw axes
		$pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Black, 1)
		$graphics.DrawLine($pen, 50, 20, 50, $chartHeight + 20) # Y-axis
		$graphics.DrawLine($pen, 50, $chartHeight + 20, $chartWidth + 50, $chartHeight + 20) # X-axis
    
		# Calculate max value for scaling
		$maxCount = 1 # Default to 1 to avoid division by zero
		foreach ($item in $sortedMonthlyPatches) {
			if ($item.Value -gt $maxCount) {
            $maxCount = $item.Value
			}
		}
    
		# Draw bars
		$barWidth = [Math]::Floor($chartWidth / ($sortedMonthlyPatches.Count + 1))
		$barX = 70
		$monthNames = @("", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
		
		# Reverse the order of the sorted monthly patches to display most recent on the right
		# First convert to array, then reverse
		$reversedPatches = @($sortedMonthlyPatches)
		[array]::Reverse($reversedPatches)
		
		
    
		foreach ($monthData in $reversedPatches) {
			$year, $month = $monthData.Name -split '-'
			$monthInt = [int]$month
			$monthName = "$($monthNames[$monthInt]) $year"
        
			$barHeight = [Math]::Round(($monthData.Value / $maxCount) * $chartHeight)
        
			# Determine bar color
			$barColor = [System.Drawing.Color]::SteelBlue
			if ("$year-$month" -eq "$currentYear-$($currentMonth.ToString('00'))") {
				$barColor = [System.Drawing.Color]::FromArgb(92, 184, 92) # Current month (green)
			}
			elseif ("$year-$month" -eq "$previousYear-$($previousMonth.ToString('00'))") {
				$barColor = [System.Drawing.Color]::FromArgb(240, 173, 78) # Previous month (yellow)
			}
        
			$brush = New-Object System.Drawing.SolidBrush($barColor)
			$graphics.FillRectangle($brush, $barX, $chartHeight + 20 - $barHeight, $barWidth, $barHeight)
			$brush.Dispose()
        
			# Draw border
			$graphics.DrawRectangle($pen, $barX, $chartHeight + 20 - $barHeight, $barWidth, $barHeight)
        
			# Draw month label
			$font = New-Object System.Drawing.Font("Arial", 8)
			$format = New-Object System.Drawing.StringFormat
			$format.Alignment = [System.Drawing.StringAlignment]::Center
			
			# Create proper rectangle objects
			$monthLabelRect = New-Object System.Drawing.RectangleF($barX, ($chartHeight + 25), $barWidth, 20)
			$graphics.DrawString($monthName, $font, [System.Drawing.Brushes]::Black, $monthLabelRect, $format)
        
			# Draw count on top of bar
			$countLabelRect = New-Object System.Drawing.RectangleF($barX, ($chartHeight + 15 - $barHeight), $barWidth, 20)
			$graphics.DrawString($monthData.Value.ToString(), $font, [System.Drawing.Brushes]::Black, $countLabelRect, $format)
        
			$barX += $barWidth + 10
			$font.Dispose()
		}
    
		$pen.Dispose()
	})

    # Create Export Report button
    $btnExportReport = New-Object System.Windows.Forms.Button
    $btnExportReport.Location = New-Object System.Drawing.Point(760, 610)
    $btnExportReport.Size = New-Object System.Drawing.Size(200, 30)
    $btnExportReport.Text = "Export Report (CSV)"
    $reportForm.Controls.Add($btnExportReport)

    # Export report event
    $btnExportReport.Add_Click({
        $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $saveFileDialog.Title = "Save Report as CSV"
        $saveFileDialog.FileName = "PatchStatusReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        
        if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $serverStatus | Export-Csv -Path $saveFileDialog.FileName -NoTypeInformation
                [System.Windows.Forms.MessageBox]::Show("Report exported successfully!", "Export Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            catch {
                [System.Windows.Forms.MessageBox]::Show("Error exporting report: $_", "Export Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    })

    # Create Close button
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(20, 610)
    $btnClose.Size = New-Object System.Drawing.Size(200, 30)
    $btnClose.Text = "Close Report"
    $reportForm.Controls.Add($btnClose)

    # Close button event
    $btnClose.Add_Click({
        $reportForm.Close()
    })

    # Show the report form
    $reportForm.ShowDialog()
}

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
        $btnReport.Enabled = $true
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

# Produce Report button click event
$btnReport.Add_Click({
    if ($script:results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No results to generate report", "No Data", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $statusLabel.Text = "Generating patch status report..."
    Show-UpdateReport -Results $script:results
    $statusLabel.Text = "Report generated"
})

# Show the form
[void]$form.ShowDialog()