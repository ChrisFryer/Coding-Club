# Force-MECMUpdates-MultiThreaded.ps1
# Multi-threaded script to trigger MECM software update installation on multiple servers
# Usage: .\Force-MECMUpdates-MultiThreaded.ps1 -CsvPath "C:\Temp\servers.csv" [-Credential $creds] [-LogPath "C:\Logs"] [-MaxThreads 10]

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter(Mandatory = $false)]
    [string]$LogPath = "C:\Temp\MECMUpdateLogs",
    
    [Parameter(Mandatory = $false)]
    [int]$MaxThreads = 10
)

# Create log directory if it doesn't exist
if (-not (Test-Path -Path $LogPath)) {
    New-Item -Path $LogPath -ItemType Directory -Force | Out-Null
}

# Log file
$LogFile = Join-Path -Path $LogPath -ChildPath "MECMUpdates_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Create a synchronized hashtable for thread-safe logging
$SyncHash = [hashtable]::Synchronized(@{})
$SyncHash.LogFile = $LogFile
$SyncHash.SuccessCount = 0
$SyncHash.FailureCount = 0
$SyncHash.TotalCount = 0
$SyncHash.Lock = New-Object System.Object

# Function to write to log in a thread-safe manner
function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $TimeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogMessage = "[$TimeStamp] [$Level] $Message"
    
    # Lock for thread safety
    [System.Threading.Monitor]::Enter($SyncHash.Lock)
    
    try {
        # Write to console with appropriate color
        switch ($Level) {
            'Info' { Write-Host $LogMessage -ForegroundColor Green }
            'Warning' { Write-Host $LogMessage -ForegroundColor Yellow }
            'Error' { Write-Host $LogMessage -ForegroundColor Red }
        }
        
        # Write to log file
        Add-Content -Path $SyncHash.LogFile -Value $LogMessage
    }
    finally {
        [System.Threading.Monitor]::Exit($SyncHash.Lock)
    }
}

Write-Log -Message "=== MECM Update Multi-Threaded Script Started ===" -Level Info
Write-Log -Message "CSV Path: $CsvPath" -Level Info
Write-Log -Message "Log Path: $LogPath" -Level Info
Write-Log -Message "Max Threads: $MaxThreads" -Level Info
Write-Log -Message "Credential Used: $(if ($Credential) { $Credential.UserName } else { 'Current User' })" -Level Info

# Import server list from CSV
try {
    $Servers = Import-Csv -Path $CsvPath
    
    # Determine which column contains server names
    $ServerNameProperty = $null
    if ($Servers[0].PSObject.Properties.Name -contains 'ServerName') {
        $ServerNameProperty = 'ServerName'
    }
    elseif ($Servers[0].PSObject.Properties.Name -contains 'ComputerName') {
        $ServerNameProperty = 'ComputerName'
    }
    elseif ($Servers[0].PSObject.Properties.Name -contains 'Name') {
        $ServerNameProperty = 'Name'
    }
    else {
        Write-Log -Message "CSV file must contain a 'ServerName', 'ComputerName', or 'Name' column" -Level Error
        exit 1
    }
    
    Write-Log -Message "Found $($Servers.Count) servers in the CSV file using column '$ServerNameProperty'" -Level Info
    $SyncHash.TotalCount = $Servers.Count
}
catch {
    Write-Log -Message "Failed to import CSV file: $($_.Exception.Message)" -Level Error
    exit 1
}

# Script block to process a single server
$ServerScriptBlock = {
    param (
        [string]$ServerName,
        [System.Management.Automation.PSCredential]$Credential,
        [hashtable]$SyncHash
    )
    
    # Function to write to log in a thread-safe manner within the runspace
    function Write-ThreadLog {
        param (
            [string]$Message,
            [string]$Level = 'Info',
            [string]$ServerName
        )
        
        $ServerPrefix = if ($ServerName) { "[$ServerName] " } else { "" }
        $TimeStamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $LogMessage = "[$TimeStamp] [$Level] $ServerPrefix$Message"
        
        # Lock for thread safety
        [System.Threading.Monitor]::Enter($SyncHash.Lock)
        
        try {
            # Write to log file directly (no console in runspace)
            Add-Content -Path $SyncHash.LogFile -Value $LogMessage
        }
        finally {
            [System.Threading.Monitor]::Exit($SyncHash.Lock)
        }
        
        # Return the message for the job output
        return $LogMessage
    }
    
    Write-ThreadLog -Message "Processing server" -ServerName $ServerName
    
    # Test if the server is online
    $pingStart = Get-Date
    $pingSuccess = Test-Connection -ComputerName $ServerName -Count 1 -Quiet
    $pingEnd = Get-Date
    $pingDuration = ($pingEnd - $pingStart).TotalMilliseconds
    
    if ($pingSuccess) {
        Write-ThreadLog -Message "Server is online (ping response: $pingDuration ms)" -ServerName $ServerName
        
        # MECM update script block to run on remote server
        $RemoteScriptBlock = {
            # Trigger software update scan cycle
            Write-Output "Initiating Software Update Scan Cycle..."
            $scanResult = Invoke-WmiMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000113}"
            Write-Output "Scan Cycle Return Value: $($scanResult.ReturnValue)"
            
            # Wait a moment for scan to process
            Write-Output "Waiting 30 seconds for scan to complete..."
            Start-Sleep -Seconds 30
            
            # Trigger software update deployment evaluation cycle
            Write-Output "Initiating Software Update Deployment Evaluation Cycle..."
            $evalResult = Invoke-WmiMethod -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000108}"
            Write-Output "Evaluation Cycle Return Value: $($evalResult.ReturnValue)"
            
            # Get current update status
            try {
                Write-Output "Checking update status..."
                $Session = New-Object -ComObject "Microsoft.CCM.UpdatesStore"
                $Updates = $Session.GetUpdates()
                
                $PendingUpdates = 0
                $DownloadingUpdates = 0
                $InstallingUpdates = 0
                $InstalledUpdates = 0
                $FailedUpdates = 0
                
                foreach ($Update in $Updates) {
                    switch ($Update.State) {
                        0 { $PendingUpdates++ }      # Available for install
                        1 { $DownloadingUpdates++ }  # Downloading
                        2 { $InstallingUpdates++ }   # Installing
                        3 { $InstalledUpdates++ }    # Installed
                        4 { $FailedUpdates++ }       # Failed
                    }
                }
                
                $statusInfo = "Update Status: Pending=$PendingUpdates, Downloading=$DownloadingUpdates, Installing=$InstallingUpdates, Installed=$InstalledUpdates, Failed=$FailedUpdates"
                Write-Output $statusInfo
                
                if ($PendingUpdates -gt 0) {
                    Write-Output "Found $PendingUpdates pending update(s). Installing now..."
                    $Updates.Install()
                    return "Installation of $PendingUpdates update(s) initiated. $statusInfo"
                }
                else {
                    return "No pending updates found. $statusInfo"
                }
            }
            catch {
                return "Error checking update status: $_"
            }
        }
        
        # Prepare parameters for Invoke-Command
        $InvokeParams = @{
            ComputerName = $ServerName
            ScriptBlock  = $RemoteScriptBlock
            ErrorAction  = 'Stop'
        }
        
        if ($Credential) {
            $InvokeParams['Credential'] = $Credential
        }
        
        # Force MECM updates
        try {
            Write-ThreadLog -Message "Executing update commands" -ServerName $ServerName
            $startTime = Get-Date
            $Result = Invoke-Command @InvokeParams
            $endTime = Get-Date
            $duration = ($endTime - $startTime).TotalSeconds
            
            Write-ThreadLog -Message "Result (completed in $duration seconds): $Result" -Level Info -ServerName $ServerName
            
            # Update success count
            [System.Threading.Monitor]::Enter($SyncHash.Lock)
            try {
                $SyncHash.SuccessCount++
            }
            finally {
                [System.Threading.Monitor]::Exit($SyncHash.Lock)
            }
            
            return @{
                ServerName = $ServerName
                Status     = "Success"
                Message    = $Result
                Duration   = $duration
            }
        }
        catch {
            Write-ThreadLog -Message "Failed to trigger updates: $($_.Exception.Message)" -Level Error -ServerName $ServerName
            
            # Try with direct WMI method as fallback
            Write-ThreadLog -Message "Attempting fallback method" -Level Warning -ServerName $ServerName
            
            try {
                $WmiParams = @{
                    ComputerName = $ServerName
                    Namespace    = "root\ccm"
                    Class        = "SMS_CLIENT"
                    Name         = "TriggerSchedule"
                    ArgumentList = "{00000000-0000-0000-0000-000000000108}"
                }
                
                if ($Credential) {
                    $WmiParams['Credential'] = $Credential
                }
                
                $wmiResult = Invoke-WmiMethod @WmiParams
                
                if ($wmiResult.ReturnValue -eq 0) {
                    Write-ThreadLog -Message "Fallback method succeeded" -Level Info -ServerName $ServerName
                    
                    # Update success count
                    [System.Threading.Monitor]::Enter($SyncHash.Lock)
                    try {
                        $SyncHash.SuccessCount++
                    }
                    finally {
                        [System.Threading.Monitor]::Exit($SyncHash.Lock)
                    }
                    
                    return @{
                        ServerName = $ServerName
                        Status     = "Success"
                        Message    = "Fallback method successful"
                        Duration   = 0
                    }
                }
                else {
                    Write-ThreadLog -Message "Fallback method failed with return code $($wmiResult.ReturnValue)" -Level Error -ServerName $ServerName
                    
                    # Update failure count
                    [System.Threading.Monitor]::Enter($SyncHash.Lock)
                    try {
                        $SyncHash.FailureCount++
                    }
                    finally {
                        [System.Threading.Monitor]::Exit($SyncHash.Lock)
                    }
                    
                    return @{
                        ServerName = $ServerName
                        Status     = "Failed"
                        Message    = "Fallback method failed with return code $($wmiResult.ReturnValue)"
                        Duration   = 0
                    }
                }
            }
            catch {
                Write-ThreadLog -Message "Fallback method failed: $($_.Exception.Message)" -Level Error -ServerName $ServerName
                
                # Update failure count
                [System.Threading.Monitor]::Enter($SyncHash.Lock)
                try {
                    $SyncHash.FailureCount++
                }
                finally {
                    [System.Threading.Monitor]::Exit($SyncHash.Lock)
                }
                
                return @{
                    ServerName = $ServerName
                    Status     = "Failed"
                    Message    = $_.Exception.Message
                    Duration   = 0
                }
            }
        }
    }
    else {
        Write-ThreadLog -Message "Server is offline or not reachable" -Level Warning -ServerName $ServerName
        
        # Update failure count
        [System.Threading.Monitor]::Enter($SyncHash.Lock)
        try {
            $SyncHash.FailureCount++
        }
        finally {
            [System.Threading.Monitor]::Exit($SyncHash.Lock)
        }
        
        return @{
            ServerName = $ServerName
            Status     = "Failed"
            Message    = "Server is offline or not reachable"
            Duration   = 0
        }
    }
}

# Setup runspace pool for multi-threading
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
$RunspacePool.Open()

$Jobs = @()
$Runspaces = New-Object System.Collections.ArrayList

Write-Log -Message "=== Starting Server Processing with $MaxThreads threads ===" -Level Info

# Start jobs for each server
foreach ($Server in $Servers) {
    $ServerName = $Server.$ServerNameProperty
    
    # Skip empty server names
    if ([string]::IsNullOrWhiteSpace($ServerName)) {
        Write-Log -Message "Skipping empty server name in CSV" -Level Warning
        continue
    }
    
    Write-Log -Message "Creating job for server: $ServerName" -Level Info
    
    # Create a PowerShell instance and add the script
    $PowerShell = [powershell]::Create().AddScript($ServerScriptBlock).AddParameter("ServerName", $ServerName).AddParameter("SyncHash", $SyncHash)
    
    if ($Credential) {
        $PowerShell = $PowerShell.AddParameter("Credential", $Credential)
    }
    
    # Use the runspace pool
    $PowerShell.RunspacePool = $RunspacePool
    
    # Start the job
    $Handle = $PowerShell.BeginInvoke()
    
    # Save job information
    $Runspace = [PSCustomObject]@{
        PowerShell = $PowerShell
        Handle     = $Handle
        Server     = $ServerName
        StartTime  = Get-Date
    }
    
    # Add to the tracking collections
    [void]$Runspaces.Add($Runspace)
}

# Monitor jobs and process results as they complete
$CompletedCount = 0
while ($CompletedCount -lt $Runspaces.Count) {
    foreach ($Runspace in $Runspaces.ToArray()) {
        if ($Runspace.Handle.IsCompleted) {
            # Get the result
            $Result = $Runspace.PowerShell.EndInvoke($Runspace.Handle)
            $Duration = ((Get-Date) - $Runspace.StartTime).TotalSeconds
            
            if ($Result) {
                $Status = $Result.Status
                $Message = $Result.Message
                Write-Log -Message "Server $($Runspace.Server) completed with status: $Status in $Duration seconds" -Level Info
            }
            else {
                Write-Log -Message "Server $($Runspace.Server) completed with no result in $Duration seconds" -Level Warning
            }
            
            # Cleanup
            $Runspace.PowerShell.Dispose()
            $Runspaces.Remove($Runspace)
            $CompletedCount++
            
            # Log progress
            $PercentComplete = [math]::Round(($CompletedCount / $SyncHash.TotalCount) * 100)
            Write-Log -Message "Progress: $CompletedCount/$($SyncHash.TotalCount) servers processed ($PercentComplete%)" -Level Info
        }
    }
    
    # Wait a bit before checking again
    Start-Sleep -Milliseconds 500
}

# Close the runspace pool
$RunspacePool.Close()
$RunspacePool.Dispose()

# Summary
Write-Log -Message "=== MECM update trigger completed ===" -Level Info
Write-Log -Message "Successful: $($SyncHash.SuccessCount), Failed: $($SyncHash.FailureCount)" -Level Info
Write-Log -Message "Log file: $($SyncHash.LogFile)" -Level Info

# Print completion message to console
$timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
Write-Host ""
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "MECM Update Trigger Complete at $timestamp" -ForegroundColor Cyan
Write-Host "Successful: $($SyncHash.SuccessCount), Failed: $($SyncHash.FailureCount)" -ForegroundColor Cyan
Write-Host "Log file: $($SyncHash.LogFile)" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan