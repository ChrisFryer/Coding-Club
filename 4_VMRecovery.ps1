# VM Upgrade Monitoring and Recovery Script
# This script silently monitors a VM after an upgrade and reboot, sending email only on failure

param(
    [Parameter(Mandatory=$true)]
    [string]$VCenter,
    
    [Parameter(Mandatory=$true)]
    [string]$VMName,
    
    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential]$Credential,
    
    [Parameter(Mandatory=$false)]
    [int]$MaxRetries = 5,
    
    [Parameter(Mandatory=$false)]
    [int]$WaitTimeMinutes = 30,
    
    [Parameter(Mandatory=$false)]
    [int]$PowerOnDelaySeconds = 60,
    
    [Parameter(Mandatory=$false)]
    [string]$SMTPServer = "smtp.company.com",
    
    [Parameter(Mandatory=$false)]
    [string]$EmailFrom = "vcenter@company.com",
    
    [Parameter(Mandatory=$false)]
    [string]$EmailTo = "admin@company.com"
)

function Send-AlertEmail {
    param(
        [string]$Subject,
        [string]$Body
    )
    
    try {
        Send-MailMessage -From $EmailFrom -To $EmailTo -Subject $Subject -Body $Body -SmtpServer $SMTPServer -UseSsl
    }
    catch {
        Write-EventLog -LogName Application -Source "VMUpgradeScript" -EntryType Error -EventId 1001 -Message "Failed to send email notification: $_" -ErrorAction SilentlyContinue
    }
}

function Test-VMConnection {
    param(
        [VMware.VimAutomation.ViCore.Types.V1.Inventory.VirtualMachine]$VM
    )
    
    # First check if VM is powered on
    if ($VM.PowerState -ne "PoweredOn") {
        return $false
    }
    
    # Check VM tools status
    $toolsStatus = (Get-View $VM.Id).Guest.ToolsRunningStatus
    if ($toolsStatus -ne "guestToolsRunning") {
        return $false
    }
    
    # Try to get VM's IP address
    $ipAddress = $VM.Guest.IPAddress | Where-Object { $_ -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}' } | Select-Object -First 1
    
    if (-not $ipAddress) {
        return $false
    }
    
    # Try to ping the VM
    $pingResult = Test-Connection -ComputerName $ipAddress -Count 2 -Quiet
    
    return $pingResult
}

# Connect to vCenter
try {
    Connect-VIServer -Server $VCenter -Credential $Credential -ErrorAction Stop | Out-Null
}
catch {
    # Silently exit on error
    exit 1
}

# Get VM object
try {
    $VM = Get-VM -Name $VMName -ErrorAction Stop
}
catch {
    # Silently disconnect and exit on error
    Disconnect-VIServer -Server $VCenter -Confirm:$false | Out-Null
    exit 1
}

# Initialize retry counter
$retryCount = 0
$success = $false

while ($retryCount -lt $MaxRetries -and -not $success) {
    # Check VM status every minute during the wait period
    $checkIntervalMinutes = 1
    $checksPerWaitPeriod = $WaitTimeMinutes / $checkIntervalMinutes
    
    for ($i = 0; $i -lt $checksPerWaitPeriod; $i++) {
        # Refresh VM object to get current state
        $VM = Get-VM -Name $VMName
        
        if (Test-VMConnection -VM $VM) {
            $success = $true
            break
        }
        
        # Wait for the check interval before trying again
        if (-not $success -and $i -lt ($checksPerWaitPeriod - 1)) {
            Start-Sleep -Seconds ($checkIntervalMinutes * 60)
        }
    }
    
    # If VM is still not available after the wait period, try recovery actions
    if (-not $success) {
        $retryCount++
        
        # Power off the VM
        Stop-VM -VM $VM -Confirm:$false -Force | Out-Null
        
        # Wait until VM is actually powered off
        $powerOffTimeout = 300 # 5 minutes
        $powerOffStart = Get-Date
        
        while ((Get-VM -Name $VMName).PowerState -ne "PoweredOff") {
            if (((Get-Date) - $powerOffStart).TotalSeconds -gt $powerOffTimeout) {
                break
            }
            
            Start-Sleep -Seconds 10
        }
        
        # Wait for specified delay before powering back on
        Start-Sleep -Seconds $PowerOnDelaySeconds
        
        # Power on the VM
        Start-VM -VM $VM | Out-Null
        
        # If this was the last retry, send failure notification
        if ($retryCount -eq $MaxRetries -and -not $success) {
            $subject = "ALERT: VM Upgrade Recovery Failed - $VMName"
            $body = @"
VM Upgrade Recovery Process Failed

VM Name: $VMName
vCenter Server: $VCenter
Maximum Retries ($MaxRetries) Reached

The VM failed to become available after multiple recovery attempts following an upgrade.
Manual intervention is required to resolve this issue.

This is an automated message from the VM Upgrade Monitoring script.
"@
            Send-AlertEmail -Subject $subject -Body $body
        }
    }
}

# Disconnect from vCenter
Disconnect-VIServer -Server $VCenter -Confirm:$false | Out-Null

# Return success status without any output
if ($success) {
    exit 0
} else {
    exit 1
}
