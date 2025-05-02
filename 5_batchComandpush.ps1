# Path to your inventory file containing server names (one per line)
$inventoryFile = ".\Documents\{file location}"

# Read server names from inventory file
$servers = Get-Content -Path $inventoryFile

# Maximum number of concurrent jobs
$maxConcurrentJobs = 5

# Store all jobs
$jobs = @()

# Prompt for credentials for cross-domain authentication
$credential = Get-Credential -Message "Enter credentials for remote server access"

# Define the script block to be executed on each server
$serverConfigScript = {
    gpupdate /force
}

# Start jobs in parallel but limit the number of concurrent jobs
foreach ($server in $servers) {
    # Wait if we've reached the maximum number of concurrent jobs
    while (($jobs | Where-Object { $_.State -eq 'Running' }).Count -ge $maxConcurrentJobs) {
        Start-Sleep -Milliseconds 500
    }
    
    # Start a new job for this server
    $job = Start-Job -ScriptBlock {
        param($serverName, $scriptBlock, $credential)
        
        try {
            # Create session with proxy bypass and specified credentials
            $session = New-PSSession -ComputerName $serverName -SessionOption (New-PSSessionOption -ProxyAccessType NoProxyServer) -Credential $using:credential -ErrorAction Stop
            
            # Execute commands on remote server
            Invoke-Command -Session $session -ScriptBlock ([ScriptBlock]::Create($scriptBlock))
            
            # Close the session
            Remove-PSSession $session
            
            # Return just the server name with "done"
            return "$serverName done"
        }
        catch {
            # Return the server name and error message
            return "$serverName failed: $($_.Exception.Message)"
        }
    } -ArgumentList $server, $serverConfigScript.ToString(), $credential
    
    # Add job to our collection
    $jobs += $job
}

# Wait for all jobs to complete
$jobs | Wait-Job

# Output all results, both successful and failed
foreach ($job in $jobs) {
    $result = Receive-Job -Job $job
    if ($result) {
        Write-Host $result
    } else {
        Write-Host "No result from job for a server"
    }
    Remove-Job -Job $job
}
