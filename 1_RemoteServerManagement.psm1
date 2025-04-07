# Remote Server Management Solution
# Author: Claude
# Date: 2025-04-02
# Description: A comprehensive PowerShell module for remote server management

#Requires -Version 5.1
#Requires -RunAsAdministrator

# Module for Remote Server Management
function New-RemoteServerSession {
    <#
    .SYNOPSIS
        Creates a new remote session to a server or list of servers.
    .DESCRIPTION
        Establishes PowerShell remote sessions to one or more servers, with options for credentials and connection parameters.
    .PARAMETER ComputerName
        The server name(s) to connect to.
    .PARAMETER Credential
        PSCredential object for authentication.
    .PARAMETER UseSSL
        Use SSL for the connection.
    .EXAMPLE
        $cred = Get-Credential
        $session = New-RemoteServerSession -ComputerName "server01", "server02" -Credential $cred -UseSSL
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,
        
        [Parameter(Mandatory = $false)]
        [switch]$UseSSL
    )
    
    begin {
        $sessions = @()
        $sessionParams = @{
            ErrorAction = "Stop"
        }
        
        if ($PSBoundParameters.ContainsKey('Credential')) {
            $sessionParams.Add('Credential', $Credential)
        }
        
        if ($UseSSL) {
            $sessionParams.Add('UseSSL', $true)
        }
    }
    
    process {
        foreach ($computer in $ComputerName) {
            try {
                Write-Verbose "Connecting to $computer..."
                $sessionParams.ComputerName = $computer
                $session = New-PSSession @sessionParams
                $sessions += $session
                Write-Verbose "Connected to $computer successfully."
            }
            catch {
                Write-Error "Failed to connect to $computer. Error: $_"
            }
        }
    }
    
    end {
        return $sessions
    }
}

function Test-RemoteServerConnection {
    <#
    .SYNOPSIS
        Tests connectivity to remote servers.
    .DESCRIPTION
        Performs connectivity tests including ping, WinRM, and optional port checks to remote servers.
    .PARAMETER ComputerName
        The server name(s) to test.
    .PARAMETER TestPorts
        Additional ports to test beyond standard WinRM ports.
    .EXAMPLE
        Test-RemoteServerConnection -ComputerName "server01", "server02" -TestPorts 80, 443
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [int[]]$TestPorts
    )
    
    begin {
        $results = @()
    }
    
    process {
        foreach ($computer in $ComputerName) {
            $result = [PSCustomObject]@{
                ComputerName = $computer
                Ping = $false
                WinRM = $false
                Ports = @{}
                Status = "Failed"
            }
            
            # Test ping
            try {
                $pingResult = Test-Connection -ComputerName $computer -Count 2 -Quiet
                $result.Ping = $pingResult
            }
            catch {
                Write-Verbose "Ping test failed for $computer. Error: $_"
            }
            
            # Test WinRM
            try {
                $winrmResult = Test-WSMan -ComputerName $computer -ErrorAction Stop
                $result.WinRM = $true
            }
            catch {
                Write-Verbose "WinRM test failed for $computer. Error: $_"
            }
            
            # Test additional ports if specified
            if ($TestPorts) {
                foreach ($port in $TestPorts) {
                    try {
                        $tcp = New-Object System.Net.Sockets.TcpClient
                        $connection = $tcp.BeginConnect($computer, $port, $null, $null)
                        $wait = $connection.AsyncWaitHandle.WaitOne(1000, $false)
                        
                        if ($wait) {
                            $tcp.EndConnect($connection)
                            $result.Ports[$port] = $true
                        }
                        else {
                            $result.Ports[$port] = $false
                        }
                        $tcp.Close()
                    }
                    catch {
                        $result.Ports[$port] = $false
                        Write-Verbose "Port $port test failed for $computer. Error: $_"
                    }
                }
            }
            
            # Determine overall status
            if ($result.Ping -and $result.WinRM) {
                $result.Status = "Ready"
            }
            elseif ($result.Ping) {
                $result.Status = "Ping Only"
            }
            
            $results += $result
        }
    }
    
    end {
        return $results
    }
}

# Main entry point for the module
function Start-RemoteServerManagement {
    <#
    .SYNOPSIS
        Main entry point for the Remote Server Management solution.
    .DESCRIPTION
        Provides an interactive menu to access all remote server management functions.
    .PARAMETER ComputerName
        The server name(s) to manage. If not specified, you will be prompted to enter them.
    .PARAMETER Credential
        PSCredential object for authentication.
    .EXAMPLE
        Start-RemoteServerManagement -ComputerName "server01", "server02"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string[]]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential
    )
    
    process {
        $title = @"
======================================================
    Remote Server Management Solution
======================================================
"@
        
        if (-not $ComputerName) {
            $ComputerName = Read-Host "Enter server name(s) to manage (comma-separated)"
            $ComputerName = $ComputerName -split ',' | ForEach-Object { $_.Trim() }
        }
        
        if (-not $Credential) {
            $Credential = Get-Credential -Message "Enter credentials for remote server access"
        }
        
        $options = @(
            "Test connectivity to servers",
            "Get server information",
            "Execute remote command",
            "Configure remote servers",
            "Monitor server performance",
            "Generate server reports",
            "Exit"
        )
        
        $running = $true
        while ($running) {
            Clear-Host
            Write-Host $title -ForegroundColor Cyan
            Write-Host "Managing servers: $($ComputerName -join ', ')" -ForegroundColor Yellow
            Write-Host ""
            
            for ($i = 0; $i -lt $options.Count; $i++) {
                Write-Host "[$($i + 1)] $($options[$i])"
            }
            
            Write-Host ""
            $choice = Read-Host "Select an option (1-$($options.Count))"
            
            switch ($choice) {
                1 {
                    # Test connectivity
                    $results = Test-RemoteServerConnection -ComputerName $ComputerName
                    
                    Clear-Host
                    Write-Host "Connectivity Test Results:" -ForegroundColor Cyan
                    $results | Format-Table -AutoSize
                    
                    $testPorts = Read-Host "Would you like to test specific ports? (y/n)"
                    if ($testPorts -eq 'y') {
                        $ports = Read-Host "Enter ports to test (comma-separated)"
                        $portArray = $ports -split ',' | ForEach-Object { [int]$_.Trim() }
                        
                        $results = Test-RemoteServerConnection -ComputerName $ComputerName -TestPorts $portArray
                        Write-Host "Port Test Results:" -ForegroundColor Cyan
                        $results | Format-Table -AutoSize
                    }
                    
                    Write-Host "Press any key to continue..."
                    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
                
                2 {
                    # Get server information
                    Clear-Host
                    Write-Host "Server Information Options:" -ForegroundColor Cyan
                    Write-Host "[1] System Information"
                    Write-Host "[2] Hardware Information"
                    Write-Host "[3] Disk Information"
                    Write-Host "[4] Network Information"
                    Write-Host "[5] Services Information"
                    Write-Host "[6] Software Information"
                    Write-Host "[7] All Information"
                    
                    $infoChoice = Read-Host "Select information to retrieve (1-7)"
                    $infoTypeMap = @{
                        1 = @("System")
                        2 = @("Hardware")
                        3 = @("Disk")
                        4 = @("Network")
                        5 = @("Services")
                        6 = @("Software")
                        7 = @("All")
                    }
                    
                    if ($infoTypeMap.ContainsKey([int]$infoChoice)) {
                        $results = Get-RemoteServerInfo -ComputerName $ComputerName -InfoType $infoTypeMap[[int]$infoChoice] -Credential $Credential
                        
                        foreach ($server in $ComputerName) {
						Write-Host "Information for ${server}:" -ForegroundColor Yellow
                            $serverResults = $results[$server]
                            
                            foreach ($category in $serverResults.Keys) {
                                Write-Host "$category Information:" -ForegroundColor Cyan
                                $serverResults[$category] | Format-List
                            }
                        }
                    }
                    else {
                        Write-Host "Invalid option selected." -ForegroundColor Red
                    }
                    
                    Write-Host "Press any key to continue..."
                    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
                
                3 {
                    # Execute remote command
                    Clear-Host
                    Write-Host "Remote Command Execution:" -ForegroundColor Cyan
                    Write-Host "[1] Execute PowerShell Command"
                    Write-Host "[2] Execute Script File"
                    
                    $cmdChoice = Read-Host "Select command type (1-2)"
                    
                    switch ($cmdChoice) {
                        1 {
                            $command = Read-Host "Enter PowerShell command to execute"
                            
                            try {
                                $scriptBlock = [scriptblock]::Create($command)
                                $results = Invoke-RemoteServerCommand -ComputerName $ComputerName -ScriptBlock $scriptBlock -Credential $Credential
                                
                                foreach ($result in $results) {
                                    Write-Host "Results from $($result.ComputerName):" -ForegroundColor Yellow
                                    Write-Host "Status: $($result.Status)" -ForegroundColor $(if ($result.Status -eq "Success") { "Green" } else { "Red" })
                                    
                                    if ($result.Status -eq "Success") {
                                        Write-Host "Output:"
                                        $result.Output
                                    }
                                    else {
                                        Write-Host "Error: $($result.Error)" -ForegroundColor Red
                                    }
                                    
                                    Write-Host "Execution Time: $($result.Duration.TotalSeconds) seconds" -ForegroundColor Cyan
                                    Write-Host ""
                                }
                            }
                            catch {
                                Write-Host "Error creating script block: $_" -ForegroundColor Red
                            }
                        }
                        
                        2 {
                            $scriptPath = Read-Host "Enter full path to script file"
                            
                            if (Test-Path $scriptPath) {
                                $results = Invoke-RemoteServerCommand -ComputerName $ComputerName -FilePath $scriptPath -Credential $Credential
                                
                                foreach ($result in $results) {
                                    Write-Host "Results from $($result.ComputerName):" -ForegroundColor Yellow
                                    Write-Host "Status: $($result.Status)" -ForegroundColor $(if ($result.Status -eq "Success") { "Green" } else { "Red" })
                                    
                                    if ($result.Status -eq "Success") {
                                        Write-Host "Output:"
                                        $result.Output
                                    }
                                    else {
                                        Write-Host "Error: $($result.Error)" -ForegroundColor Red
                                    }
                                    
                                    Write-Host "Execution Time: $($result.Duration.TotalSeconds) seconds" -ForegroundColor Cyan
                                    Write-Host ""
                                }
                            }
                            else {
                                Write-Host "Script file not found: $scriptPath" -ForegroundColor Red
                            }
                        }
                        
                        default {
                            Write-Host "Invalid option selected." -ForegroundColor Red
                        }
                    }
                    
                    Write-Host "Press any key to continue..."
                    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
                
                4 {
                    # Configure remote servers
                    Clear-Host
                    Write-Host "Remote Server Configuration:" -ForegroundColor Cyan
                    Write-Host "[1] Configure Services"
                    Write-Host "[2] Configure Scheduled Tasks"
                    Write-Host "[3] Configure Registry Settings"
                    Write-Host "[4] Configure Windows Features"
                    Write-Host "[5] Configure Firewall Rules"
                    
                    $configChoice = Read-Host "Select configuration type (1-5)"
                    
                    $configTypeMap = @{
                        1 = "Service"
                        2 = "ScheduledTask"
                        3 = "Registry"
                        4 = "WindowsFeature"
                        5 = "FirewallRule"
                    }
                    
                    if ($configTypeMap.ContainsKey([int]$configChoice)) {
                        $configType = $configTypeMap[[int]$configChoice]
                        
                        switch ($configType) {
                            "Service" {
                                $serviceName = Read-Host "Enter service name"
                                $serviceStatus = Read-Host "Enter desired status (Running/Stopped)"
                                $startupType = Read-Host "Enter startup type (Automatic/Manual/Disabled)"
                                
                                $configParams = @{
                                    Name = $serviceName
                                    Status = $serviceStatus
                                    StartupType = $startupType
                                }
                            }
                            
                            "ScheduledTask" {
                                $taskName = Read-Host "Enter task name"
                                $taskAction = Read-Host "Enter action (Enable/Disable/Run)"
                                
                                $configParams = @{
                                    TaskName = $taskName
                                    Action = $taskAction
                                }
                            }
                            
                            "Registry" {
                                $regPath = Read-Host "Enter registry path (e.g., HKLM:\SOFTWARE\...)"
                                $regName = Read-Host "Enter value name"
                                $regValue = Read-Host "Enter value data"
                                $regType = Read-Host "Enter value type (String/DWord/Binary/...)"
                                
                                $configParams = @{
                                    Path = $regPath
                                    Name = $regName
                                    Value = $regValue
                                    Type = $regType
                                }
                            }
                            
                            "WindowsFeature" {
                                $featureName = Read-Host "Enter Windows feature name"
                                $featureAction = Read-Host "Enter action (Install/Uninstall)"
                                
                                $configParams = @{
                                    Name = $featureName
                                    Action = $featureAction
                                }
                            }
                            
                            "FirewallRule" {
                                $ruleName = Read-Host "Enter firewall rule name"
                                $ruleAction = Read-Host "Enter action (Enable/Disable/Create)"
                                
                                $configParams = @{
                                    Name = $ruleName
                                    Action = $ruleAction
                                }
                                
                                if ($ruleAction -eq "Create") {
                                    $displayName = Read-Host "Enter display name"
                                    $direction = Read-Host "Enter direction (Inbound/Outbound)"
                                    $enabled = Read-Host "Enter enabled state (True/False)"
                                    $action = Read-Host "Enter action (Allow/Block)"
                                    $protocol = Read-Host "Enter protocol (TCP/UDP/Any)"
                                    $localPort = Read-Host "Enter local port(s)"
                                    
                                    $configParams["DisplayName"] = $displayName
                                    $configParams["Direction"] = $direction
                                    $configParams["Enabled"] = [bool]::Parse($enabled)
                                    $configParams["Action"] = $action
                                    $configParams["Protocol"] = $protocol
                                    
                                    if ($localPort -ne "") {
                                        $configParams["LocalPort"] = $localPort
                                    }
                                }
                            }
                        }
                        
                        $confirm = Read-Host "Apply this configuration to all servers? (y/n)"
                        if ($confirm -eq "y") {
                            $results = Set-RemoteServerConfiguration -ComputerName $ComputerName -ConfigType $configType -ConfigParams $configParams -Credential $Credential
                            
                            foreach ($result in $results) {
                                Write-Host "Results from $($result.ComputerName):" -ForegroundColor Yellow
                                Write-Host "Status: $($result.Status)" -ForegroundColor $(if ($result.Status -eq "Success") { "Green" } else { "Red" })
                                
                                if ($result.Status -eq "Success") {
                                    Write-Host "Output:"
                                    $result.Output
                                }
                                else {
                                    Write-Host "Error: $($result.Error)" -ForegroundColor Red
                                }
                                
                                Write-Host ""
                            }
                        }
                    }
                    else {
                        Write-Host "Invalid option selected." -ForegroundColor Red
                    }
                    
                    Write-Host "Press any key to continue..."
                    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
                
                5 {
                    # Monitor server performance
                    Clear-Host
                    Write-Host "Performance Monitoring:" -ForegroundColor Cyan
                    Write-Host "[1] CPU Metrics"
                    Write-Host "[2] Memory Metrics"
                    Write-Host "[3] Disk Metrics"
                    Write-Host "[4] Network Metrics"
                    Write-Host "[5] All Metrics"
                    
                    $perfChoice = Read-Host "Select metrics to monitor (1-5)"
                    $metricTypeMap = @{
                        1 = @("CPU")
                        2 = @("Memory")
                        3 = @("Disk")
                        4 = @("Network")
                        5 = @("All")
                    }
                    
                    if ($metricTypeMap.ContainsKey([int]$perfChoice)) {
                        $sampleInterval = Read-Host "Enter sample interval in seconds (default: 5)"
                        if ([string]::IsNullOrEmpty($sampleInterval)) { $sampleInterval = 5 }
                        
                        $sampleCount = Read-Host "Enter number of samples to collect (default: 3)"
                        if ([string]::IsNullOrEmpty($sampleCount)) { $sampleCount = 3 }
                        
                        Write-Host "Collecting performance metrics..." -ForegroundColor Yellow
                        $results = Get-RemoteServerPerformance -ComputerName $ComputerName -MetricType $metricTypeMap[[int]$perfChoice] -SampleInterval ([int]$sampleInterval) -SampleCount ([int]$sampleCount) -Credential $Credential
                        
                        foreach ($server in $ComputerName) {
                            Write-Host "Performance results for ${server}:" -ForegroundColor Yellow
                            $serverResults = $results[$server]
                            
                            if ($serverResults -is [string]) {
                                Write-Host "Error: $serverResults" -ForegroundColor Red
                                continue
                            }
                            
                            foreach ($category in $serverResults.Keys) {
                                Write-Host "$category Metrics:" -ForegroundColor Cyan
                                
                                $metrics = @()
                                foreach ($counterKey in $serverResults[$category].Keys) {
                                    $counterSamples = $serverResults[$category][$counterKey]
                                    $avgValue = ($counterSamples | Measure-Object -Property Value -Average).Average
                                    
                                    $metrics += [PSCustomObject]@{
                                        Counter = $counterSamples[0].Counter
                                        Instance = $counterSamples[0].Instance
                                        AverageValue = [math]::Round($avgValue, 2)
                                    }
                                }
                                
                                $metrics | Format-Table -AutoSize
                            }
                        }
                    }
                    else {
                        Write-Host "Invalid option selected." -ForegroundColor Red
                    }
                    
                    Write-Host "Press any key to continue..."
                    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
                
                6 {
                    # Generate server reports
                    Clear-Host
                    Write-Host "Server Reporting:" -ForegroundColor Cyan
                    Write-Host "[1] Generate HTML Report"
                    Write-Host "[2] Generate CSV Report"
                    Write-Host "[3] Generate XML Report"
                    
                    $reportChoice = Read-Host "Select report type (1-3)"
                    $reportTypeMap = @{
                        1 = "HTML"
                        2 = "CSV"
                        3 = "XML"
                    }
                    
                    if ($reportTypeMap.ContainsKey([int]$reportChoice)) {
                        $reportType = $reportTypeMap[[int]$reportChoice]
                        
                        Write-Host "Select report content:" -ForegroundColor Cyan
                        Write-Host "[1] Server Information"
                        Write-Host "[2] Performance Metrics"
                        Write-Host "[3] Services"
                        Write-Host "[4] System Inventory"
                        Write-Host "[5] All of the above"
                        
                        $contentChoice = Read-Host "Select content to include (comma-separated, e.g., 1,3,4)"
                        $contentOptions = $contentChoice -split ',' | ForEach-Object { $_.Trim() }
                        
                        $contentTypeMap = @{
                            1 = "ServerInfo"
                            2 = "Performance"
                            3 = "Services"
                            4 = "Inventory"
                            5 = "All"
                        }
                        
                        $reportContent = @()
                        foreach ($option in $contentOptions) {
                            if ($contentTypeMap.ContainsKey([int]$option)) {
                                $reportContent += $contentTypeMap[[int]$option]
                            }
                        }
                        
                        if ($reportContent.Count -eq 0) {
                            Write-Host "No valid content options selected." -ForegroundColor Red
                        }
                        else {
                            $outputPath = Read-Host "Enter output directory path"
                            
                            if (-not (Test-Path -Path $outputPath)) {
                                New-Item -Path $outputPath -ItemType Directory -Force | Out-Null
                            }
                            
                            Write-Host "Generating reports..." -ForegroundColor Yellow
                            Export-RemoteServerReport -ComputerName $ComputerName -ReportType $reportType -OutputPath $outputPath -ReportContent $reportContent -Credential $Credential
                            
                            Write-Host "Reports generated successfully in $outputPath" -ForegroundColor Green
                        }
                    }
                    else {
                        Write-Host "Invalid option selected." -ForegroundColor Red
                    }
                    
                    Write-Host "Press any key to continue..."
                    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
                
                7 {
                    # Exit
                    $running = $false
                    Write-Host "Exiting Remote Server Management..." -ForegroundColor Yellow
                }
                
                default {
                    Write-Host "Invalid option selected." -ForegroundColor Red
                    Write-Host "Press any key to continue..."
                    $null = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                }
            }
        }
    }
}

# Export module functions
Export-ModuleMember -Function @(
    'New-RemoteServerSession',
    'Test-RemoteServerConnection',
    'Get-RemoteServerInfo',
    'Invoke-RemoteServerCommand',
    'Set-RemoteServerConfiguration',
    'Get-RemoteServerPerformance',
    'Export-RemoteServerReport',
    'Start-RemoteServerManagement'
)

function Get-RemoteServerInfo {
    <#
    .SYNOPSIS
        Retrieves detailed system information from remote servers.
    .DESCRIPTION
        Collects comprehensive system information including OS details, hardware, services, and more from remote servers.
    .PARAMETER ComputerName
        The server name(s) to query.
    .PARAMETER Credential
        PSCredential object for authentication.
    .PARAMETER InfoType
        Type of information to retrieve (System, Hardware, Disk, Network, Services, Software, All).
    .EXAMPLE
        Get-RemoteServerInfo -ComputerName "server01" -InfoType System,Disk -Credential (Get-Credential)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("System", "Hardware", "Disk", "Network", "Services", "Software", "All")]
        [string[]]$InfoType = @("System")
    )
    
    begin {
        $results = @{}
        $scriptBlocks = @{
            System = {
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    OSVersion = (Get-CimInstance Win32_OperatingSystem).Caption
                    OSBuild = (Get-CimInstance Win32_OperatingSystem).BuildNumber
                    LastBoot = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
                    Uptime = (Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
                    InstalledRam = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
                    PSVersion = $PSVersionTable.PSVersion.ToString()
                }
            }
            Hardware = {
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    Manufacturer = (Get-CimInstance Win32_ComputerSystem).Manufacturer
                    Model = (Get-CimInstance Win32_ComputerSystem).Model
                    SerialNumber = (Get-CimInstance Win32_BIOS).SerialNumber
                    Processors = @(Get-CimInstance Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors)
                    PhysicalMemory = @(Get-CimInstance Win32_PhysicalMemory | Select-Object DeviceLocator, Capacity, Speed)
                }
            }
            Disk = {
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    Volumes = @(Get-Volume | Where-Object { $_.DriveLetter } | Select-Object DriveLetter, FileSystemLabel, FileSystem, 
                              @{Name="SizeGB"; Expression={[math]::Round(($_.Size / 1GB), 2)}},
                              @{Name="FreeSpaceGB"; Expression={[math]::Round(($_.SizeRemaining / 1GB), 2)}},
                              @{Name="FreePercent"; Expression={[math]::Round(($_.SizeRemaining / $_.Size * 100), 2)}})
                    PhysicalDisks = @(Get-PhysicalDisk | Select-Object FriendlyName, MediaType, OperationalStatus, HealthStatus, 
                                    @{Name="SizeGB"; Expression={[math]::Round(($_.Size / 1GB), 2)}})
                }
            }
            Network = {
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    IPConfiguration = @(Get-NetIPConfiguration | Select-Object InterfaceAlias, InterfaceDescription, IPv4Address, IPv4DefaultGateway)
                    DNSServers = @(Get-DnsClientServerAddress | Where-Object { $_.AddressFamily -eq 2 } | Select-Object InterfaceAlias, ServerAddresses)
                    OpenPorts = @(Get-NetTCPConnection | Where-Object { $_.State -eq "Listen" } | Group-Object LocalPort | Select-Object Name, Count)
                }
            }
            Services = {
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    RunningServices = @(Get-Service | Where-Object { $_.Status -eq "Running" } | Select-Object Name, DisplayName, StartType)
                    StoppedServices = @(Get-Service | Where-Object { $_.Status -eq "Stopped" -and $_.StartType -ne "Disabled" } | Select-Object Name, DisplayName, StartType)
                    AutoStartServices = @(Get-Service | Where-Object { $_.StartType -eq "Automatic" } | Select-Object Name, DisplayName, Status)
                }
            }
            Software = {
                [PSCustomObject]@{
                    ComputerName = $env:COMPUTERNAME
                    InstalledSoftware = @(Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
                                        Where-Object { $_.DisplayName } | 
                                        Select-Object DisplayName, DisplayVersion, Publisher, InstallDate)
                    InstalledHotfixes = @(Get-HotFix | Select-Object HotFixID, Description, InstalledOn | Sort-Object InstalledOn -Descending)
                    PendingReboot = (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending") -or
                                  (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
                }
            }
        }

        if ($InfoType -contains "All") {
            $InfoType = @("System", "Hardware", "Disk", "Network", "Services", "Software")
        }
        
        $invocationParams = @{
            ErrorAction = "Stop"
        }
        
        if ($PSBoundParameters.ContainsKey('Credential')) {
            $invocationParams.Add('Credential', $Credential)
        }
    }
    
    process {
        foreach ($computer in $ComputerName) {
            $results[$computer] = @{}
            
            foreach ($type in $InfoType) {
                try {
                    Write-Verbose "Retrieving $type information from $computer..."
                    $invocationParams.ComputerName = $computer
                    $invocationParams.ScriptBlock = $scriptBlocks[$type]
                    
                    $info = Invoke-Command @invocationParams
                    $results[$computer][$type] = $info
                    Write-Verbose "Retrieved $type information from $computer successfully."
                }
                catch {
                    Write-Error "Failed to retrieve $type information from $computer. Error: $_"
                    $results[$computer][$type] = "Error: $_"
                }
            }
        }
    }
    
    end {
        return $results
    }
}

function Invoke-RemoteServerCommand {
    <#
    .SYNOPSIS
        Executes commands on remote servers.
    .DESCRIPTION
        Runs PowerShell commands or scripts on one or more remote servers with detailed output.
    .PARAMETER ComputerName
        The server name(s) to run commands on.
    .PARAMETER ScriptBlock
        The PowerShell script block to execute.
    .PARAMETER FilePath
        Path to a script file to execute.
    .PARAMETER ArgumentList
        Arguments to pass to the script block or file.
    .PARAMETER Credential
        PSCredential object for authentication.
    .EXAMPLE
        Invoke-RemoteServerCommand -ComputerName "server01", "server02" -ScriptBlock { Get-Process -Name "*sql*" } -Credential (Get-Credential)
    #>
    [CmdletBinding(DefaultParameterSetName = "ScriptBlock")]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$ComputerName,
        
        [Parameter(Mandatory = $true, ParameterSetName = "ScriptBlock")]
        [scriptblock]$ScriptBlock,
        
        [Parameter(Mandatory = $true, ParameterSetName = "FilePath")]
        [string]$FilePath,
        
        [Parameter(Mandatory = $false)]
        [object[]]$ArgumentList,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    )
    
    begin {
        $results = @()
        $invocationParams = @{
            ErrorAction = "Stop"
        }
        
        if ($PSBoundParameters.ContainsKey('Credential')) {
            $invocationParams.Add('Credential', $Credential)
        }
        
        if ($PSBoundParameters.ContainsKey('ArgumentList')) {
            $invocationParams.Add('ArgumentList', $ArgumentList)
        }
        
        if ($PSCmdlet.ParameterSetName -eq "FilePath") {
            if (-not (Test-Path $FilePath)) {
                throw "Script file not found: $FilePath"
            }
            $scriptContent = Get-Content -Path $FilePath -Raw
            $ScriptBlock = [scriptblock]::Create($scriptContent)
        }
        
        $invocationParams.Add('ScriptBlock', $ScriptBlock)
    }
    
    process {
        foreach ($computer in $ComputerName) {
            $result = [PSCustomObject]@{
                ComputerName = $computer
                Status = "Failed"
                Output = $null
                Error = $null
                StartTime = Get-Date
                EndTime = $null
                Duration = $null
            }
            
            try {
                Write-Verbose "Executing command on $computer..."
                $invocationParams.ComputerName = $computer
                
                $output = Invoke-Command @invocationParams
                $result.Output = $output
                $result.Status = "Success"
                Write-Verbose "Command executed successfully on $computer."
            }
            catch {
                Write-Error "Failed to execute command on $computer. Error: $_"
                $result.Error = $_
            }
            finally {
                $result.EndTime = Get-Date
                $result.Duration = $result.EndTime - $result.StartTime
                $results += $result
            }
        }
    }
    
    end {
        return $results
    }
}

function Set-RemoteServerConfiguration {
    <#
    .SYNOPSIS
        Applies configuration changes to remote servers.
    .DESCRIPTION
        Configures various aspects of remote servers including services, scheduled tasks, registry, and more.
    .PARAMETER ComputerName
        The server name(s) to configure.
    .PARAMETER ConfigType
        Type of configuration to apply (Service, ScheduledTask, Registry, WindowsFeature).
    .PARAMETER ConfigParams
        Hashtable of configuration parameters specific to the ConfigType.
    .PARAMETER Credential
        PSCredential object for authentication.
    .EXAMPLE
        Set-RemoteServerConfiguration -ComputerName "server01" -ConfigType Service -ConfigParams @{Name="Spooler"; Status="Running"; StartupType="Automatic"} -Credential (Get-Credential)
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("Service", "ScheduledTask", "Registry", "WindowsFeature", "FirewallRule")]
        [string]$ConfigType,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$ConfigParams,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    )
    
    begin {
        $results = @()
        $invocationParams = @{
            ErrorAction = "Stop"
        }
        
        if ($PSBoundParameters.ContainsKey('Credential')) {
            $invocationParams.Add('Credential', $Credential)
        }
        
        # Define script blocks for each configuration type
        $scriptBlocks = @{
            Service = {
                param($params)
                $serviceName = $params.Name
                $status = $params.Status
                $startType = $params.StartupType
                
                $service = Get-Service -Name $serviceName -ErrorAction Stop
                
                if ($startType) {
                    Set-Service -Name $serviceName -StartupType $startType
                }
                
                if ($status -eq "Running" -and $service.Status -ne "Running") {
                    Start-Service -Name $serviceName
                }
                elseif ($status -eq "Stopped" -and $service.Status -ne "Stopped") {
                    Stop-Service -Name $serviceName
                }
                
                Get-Service -Name $serviceName
            }
            
            ScheduledTask = {
                param($params)
                $taskName = $params.TaskName
                $taskAction = $params.Action
                
                if ($taskAction -eq "Enable") {
                    Enable-ScheduledTask -TaskName $taskName
                }
                elseif ($taskAction -eq "Disable") {
                    Disable-ScheduledTask -TaskName $taskName
                }
                elseif ($taskAction -eq "Run") {
                    Start-ScheduledTask -TaskName $taskName
                }
                elseif ($taskAction -eq "Create") {
                    $trigger = New-ScheduledTaskTrigger -At $params.Trigger
					$action = New-ScheduledTaskAction -Execute $params.Action
					$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries $params.Settings
                    
                    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
                        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                    }
                    
                    Register-ScheduledTask -TaskName $taskName -Trigger $trigger -Action $action -Settings $settings
                }
                
                Get-ScheduledTask -TaskName $taskName
            }
            
            Registry = {
                param($params)
                $path = $params.Path
                $name = $params.Name
                $value = $params.Value
                $type = $params.Type
                
                if (-not (Test-Path $path)) {
                    New-Item -Path $path -Force | Out-Null
                }
                
                Set-ItemProperty -Path $path -Name $name -Value $value -Type $type
                
                Get-ItemProperty -Path $path -Name $name
            }
            
            WindowsFeature = {
                param($params)
                $featureName = $params.Name
                $action = $params.Action
                
                if ($action -eq "Install") {
                    Install-WindowsFeature -Name $featureName -IncludeManagementTools
                }
                elseif ($action -eq "Uninstall") {
                    Uninstall-WindowsFeature -Name $featureName
                }
                
                Get-WindowsFeature -Name $featureName
            }
            
            FirewallRule = {
                param($params)
                $ruleName = $params.Name
                $action = $params.Action
                
                if ($action -eq "Enable") {
                    Enable-NetFirewallRule -Name $ruleName
                }
                elseif ($action -eq "Disable") {
                    Disable-NetFirewallRule -Name $ruleName
                }
                elseif ($action -eq "Create") {
                    $ruleParams = $params.Clone()
                    $ruleParams.Remove('Action')
                    
                    if (Get-NetFirewallRule -Name $ruleName -ErrorAction SilentlyContinue) {
                        Remove-NetFirewallRule -Name $ruleName
                    }
                    
                    New-NetFirewallRule @ruleParams
                }
                
                Get-NetFirewallRule -Name $ruleName
            }
        }
    }
    
    process {
        foreach ($computer in $ComputerName) {
            $result = [PSCustomObject]@{
                ComputerName = $computer
                ConfigType = $ConfigType
                Status = "Failed"
                Output = $null
                Error = $null
            }
            
            if ($PSCmdlet.ShouldProcess($computer, "Apply $ConfigType configuration")) {
                try {
                    Write-Verbose "Applying $ConfigType configuration to $computer..."
                    $invocationParams.ComputerName = $computer
                    $invocationParams.ScriptBlock = $scriptBlocks[$ConfigType]
                    $invocationParams.ArgumentList = $ConfigParams
                    
                    $output = Invoke-Command @invocationParams
                    $result.Output = $output
                    $result.Status = "Success"
                    Write-Verbose "Applied $ConfigType configuration to $computer successfully."
                }
                catch {
                    Write-Error "Failed to apply $ConfigType configuration to $computer. Error: $_"
                    $result.Error = $_
                }
                finally {
                    $results += $result
                }
            }
        }
    }
    
    end {
        return $results
    }
}

function Get-RemoteServerPerformance {
    <#
    .SYNOPSIS
        Monitors performance metrics on remote servers.
    .DESCRIPTION
        Collects and reports performance metrics including CPU, memory, disk, and network utilization on remote servers.
    .PARAMETER ComputerName
        The server name(s) to monitor.
    .PARAMETER MetricType
        Type of metrics to collect (CPU, Memory, Disk, Network, All).
    .PARAMETER SampleInterval
        Interval in seconds between samples.
    .PARAMETER SampleCount
        Number of samples to collect.
    .PARAMETER Credential
        PSCredential object for authentication.
    .EXAMPLE
        Get-RemoteServerPerformance -ComputerName "server01" -MetricType CPU,Memory -SampleInterval 5 -SampleCount 10
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string[]]$ComputerName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("CPU", "Memory", "Disk", "Network", "All")]
        [string[]]$MetricType = @("CPU", "Memory"),
        
        [Parameter(Mandatory = $false)]
        [int]$SampleInterval = 5,
        
        [Parameter(Mandatory = $false)]
        [int]$SampleCount = 1,
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    )
    
    begin {
        $results = @{}
        
        if ($MetricType -contains "All") {
            $MetricType = @("CPU", "Memory", "Disk", "Network")
        }
        
        $counters = @{
            CPU = @(
                "\Processor(_Total)\% Processor Time",
                "\System\Processor Queue Length"
            )
            Memory = @(
                "\Memory\Available MBytes",
                "\Memory\% Committed Bytes In Use",
                "\Memory\Pages/sec"
            )
            Disk = @(
                "\LogicalDisk(*)\% Disk Time",
                "\LogicalDisk(*)\Avg. Disk Queue Length",
                "\LogicalDisk(*)\Disk Reads/sec",
                "\LogicalDisk(*)\Disk Writes/sec"
            )
            Network = @(
                "\Network Interface(*)\Bytes Total/sec",
                "\Network Interface(*)\Current Bandwidth",
                "\TCPv4\Connections Established"
            )
        }
        
        $selectedCounters = @()
        foreach ($metric in $MetricType) {
            $selectedCounters += $counters[$metric]
        }
        
        $invocationParams = @{
            ErrorAction = "Stop"
        }
        
        if ($PSBoundParameters.ContainsKey('Credential')) {
            $invocationParams.Add('Credential', $Credential)
        }
    }
    
    process {
        foreach ($computer in $ComputerName) {
            $results[$computer] = @{}
            
            try {
                Write-Verbose "Collecting performance metrics from $computer..."
                $invocationParams.ComputerName = $computer
                $invocationParams.ScriptBlock = {
                    param($counters, $interval, $count)
                    
                    $samples = Get-Counter -Counter $counters -SampleInterval $interval -MaxSamples $count -ErrorAction Stop
                    return $samples
                }
                $invocationParams.ArgumentList = $selectedCounters, $SampleInterval, $SampleCount
                
                $perfData = Invoke-Command @invocationParams
                
                # Process and organize the performance data
                foreach ($sample in $perfData) {
                    $timestamp = $sample.Timestamp
                    
                    foreach ($counterResult in $sample.CounterSamples) {
                        $counterPath = $counterResult.Path
                        $counterValue = $counterResult.CookedValue
                        
                        $metricCategory = "Other"
                        foreach ($metric in $MetricType) {
                            if ($counterPath -match $metric) {
                                $metricCategory = $metric
                                break
                            }
                        }
                        
                        if (-not $results[$computer].ContainsKey($metricCategory)) {
                            $results[$computer][$metricCategory] = @{}
                        }
                        
                        $counterName = ($counterPath -split '\\')[-1]
                        $instanceName = if ($counterPath -match '\((.*?)\)') { $matches[1] } else { "_Total" }
                        
                        $key = "$counterName|$instanceName"
                        
                        if (-not $results[$computer][$metricCategory].ContainsKey($key)) {
                            $results[$computer][$metricCategory][$key] = @()
                        }
                        
                        $results[$computer][$metricCategory][$key] += [PSCustomObject]@{
                            Timestamp = $timestamp
                            Value = $counterValue
                            Instance = $instanceName
                            Counter = $counterName
                        }
                    }
                }
                
                Write-Verbose "Collected performance metrics from $computer successfully."
            }
            catch {
                Write-Error "Failed to collect performance metrics from $computer. Error: $_"
                $results[$computer] = "Error: $_"
            }
        }
    }
    
    end {
        return $results
    }
}

function Export-RemoteServerReport {
    <#
    .SYNOPSIS
        Generates comprehensive reports for remote servers.
    .DESCRIPTION
        Creates detailed HTML, CSV, or XML reports of server information, performance, and configuration.
    .PARAMETER ComputerName
        The server name(s) to report on.
    .PARAMETER ReportType
        Type of report to generate (HTML, CSV, XML).
    .PARAMETER OutputPath
        Path where the report will be saved.
    .PARAMETER ReportContent
        Content to include in the report (ServerInfo, Performance, Services, Inventory, All).
    .PARAMETER Credential
        PSCredential object for authentication.
    .EXAMPLE
        Export-RemoteServerReport -ComputerName "server01", "server02" -ReportType HTML -OutputPath C:\Reports -ReportContent All
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ComputerName,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet("HTML", "CSV", "XML")]
        [string]$ReportType,
        
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("ServerInfo", "Performance", "Services", "Inventory", "All")]
        [string[]]$ReportContent = @("ServerInfo"),
        
        [Parameter(Mandatory = $false)]
        [System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
    )
    
    begin {
        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }
        
        if ($ReportContent -contains "All") {
            $ReportContent = @("ServerInfo", "Performance", "Services", "Inventory")
        }
        
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $reportData = @{}
        
        $getServerInfoParams = @{
            ComputerName = $ComputerName
        }
        
        if ($PSBoundParameters.ContainsKey('Credential')) {
            $getServerInfoParams.Add('Credential', $Credential)
        }
    }
    
    process {
        # Collect data for the report
        if ($ReportContent -contains "ServerInfo") {
            Write-Verbose "Collecting server information..."
            $getServerInfoParams.InfoType = @("System", "Hardware")
            $reportData.ServerInfo = Get-RemoteServerInfo @getServerInfoParams
        }
        
        if ($ReportContent -contains "Performance") {
            Write-Verbose "Collecting performance metrics..."
            $perfParams = @{
                ComputerName = $ComputerName
                MetricType = @("CPU", "Memory", "Disk")
                SampleInterval = 5
                SampleCount = 3
            }
            
            if ($PSBoundParameters.ContainsKey('Credential')) {
                $perfParams.Add('Credential', $Credential)
            }
            
            $reportData.Performance = Get-RemoteServerPerformance @perfParams
        }
        
        if ($ReportContent -contains "Services") {
            Write-Verbose "Collecting service information..."
            $getServerInfoParams.InfoType = @("Services")
            $reportData.Services = Get-RemoteServerInfo @getServerInfoParams
        }
        
        if ($ReportContent -contains "Inventory") {
            Write-Verbose "Collecting inventory information..."
            $getServerInfoParams.InfoType = @("Software", "Disk", "Network")
            $reportData.Inventory = Get-RemoteServerInfo @getServerInfoParams
        }
        
        # Generate the report based on the specified type
        foreach ($computer in $ComputerName) {
            $reportFilePath = Join-Path -Path $OutputPath -ChildPath "ServerReport_${computer}_${timestamp}"
            
            switch ($ReportType) {
                "HTML" {
                    $reportFilePath += ".html"
                    Write-Verbose "Generating HTML report for $computer..."
                    
                    # Create HTML header
                    $htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <title>Server Report - $computer</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1, h2, h3 { color: #0066cc; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th { background-color: #0066cc; color: white; text-align: left; padding: 8px; }
        td { border: 1px solid #ddd; padding: 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .section { margin-bottom: 30px; }
        .summary { background-color: #f8f8f8; padding: 10px; border-left: 5px solid #0066cc; }
    </style>
</head>
<body>
    <h1>Server Report - $computer</h1>
    <div class="summary">
        <p>Report generated on $(Get-Date)</p>
    </div>
"@
                    
                    # Add server information section if requested
                    if ($ReportContent -contains "ServerInfo" -and $reportData.ServerInfo[$computer]) {
                        $htmlReport += @"
    <div class="section">
        <h2>System Information</h2>
        <table>
            <tr>
                <th>Property</th>
                <th>Value</th>
            </tr>
"@
                        
                        $sysInfo = $reportData.ServerInfo[$computer]["System"]
                        if ($sysInfo -and $sysInfo -isnot [string]) {
                            $htmlReport += @"
            <tr>
                <td>OS Version</td>
                <td>$($sysInfo.OSVersion)</td>
            </tr>
            <tr>
                <td>OS Build</td>
                <td>$($sysInfo.OSBuild)</td>
            </tr>
            <tr>
                <td>Last Boot</td>
                <td>$($sysInfo.LastBoot)</td>
            </tr>
            <tr>
                <td>Uptime</td>
                <td>$($sysInfo.Uptime)</td>
            </tr>
            <tr>
                <td>Installed RAM (GB)</td>
                <td>$($sysInfo.InstalledRam)</td>
            </tr>
            <tr>
                <td>PowerShell Version</td>
                <td>$($sysInfo.PSVersion)</td>
            </tr>
"@
                        }

                        $hwInfo = $reportData.ServerInfo[$computer]["Hardware"]
                        if ($hwInfo -and $hwInfo -isnot [string]) {
                            $htmlReport += @"
            <tr>
                <td>Manufacturer</td>
                <td>$($hwInfo.Manufacturer)</td>
            </tr>
            <tr>
                <td>Model</td>
                <td>$($hwInfo.Model)</td>
            </tr>
            <tr>
                <td>Serial Number</td>
                <td>$($hwInfo.SerialNumber)</td>
            </tr>
"@
                        }
                        
                        $htmlReport += @"
        </table>
    </div>
"@
                    }
                    
                    # Add performance section if requested
                    if ($ReportContent -contains "Performance" -and $reportData.Performance[$computer]) {
                        $htmlReport += @"
    <div class="section">
        <h2>Performance Metrics</h2>
"@
                        
                        $perfData = $reportData.Performance[$computer]
                        if ($perfData -and $perfData -isnot [string]) {
                            foreach ($metricCategory in $perfData.Keys) {
                                $htmlReport += @"
        <h3>$metricCategory Metrics</h3>
        <table>
            <tr>
                <th>Counter</th>
                <th>Instance</th>
                <th>Average Value</th>
            </tr>
"@
                                
                                foreach ($counterKey in $perfData[$metricCategory].Keys) {
                                    $counterSamples = $perfData[$metricCategory][$counterKey]
                                    $avgValue = ($counterSamples | Measure-Object -Property Value -Average).Average
                                    $counter = $counterSamples[0].Counter
                                    $instance = $counterSamples[0].Instance
                                    
                                    $htmlReport += @"
            <tr>
                <td>$counter</td>
                <td>$instance</td>
                <td>$([math]::Round($avgValue, 2))</td>
            </tr>
"@
                                }
                                
                                $htmlReport += @"
        </table>
"@
                            }
                        }
                        
                        $htmlReport += @"
    </div>
"@
                    }
                    
                    # Add services section if requested
                    if ($ReportContent -contains "Services" -and $reportData.Services[$computer]) {
                        $htmlReport += @"
    <div class="section">
        <h2>Services</h2>
"@
                        
                        $servicesData = $reportData.Services[$computer]["Services"]
                        if ($servicesData -and $servicesData -isnot [string]) {
                            $htmlReport += @"
        <h3>Running Services</h3>
        <table>
            <tr>
                <th>Name</th>
                <th>Display Name</th>
                <th>Start Type</th>
            </tr>
"@
                            
                            foreach ($service in $servicesData.RunningServices) {
                                $htmlReport += @"
            <tr>
                <td>$($service.Name)</td>
                <td>$($service.DisplayName)</td>
                <td>$($service.StartType)</td>
            </tr>
"@
                            }
                            
                            $htmlReport += @"
        </table>
        
        <h3>Stopped Services (Non-Disabled)</h3>
        <table>
            <tr>
                <th>Name</th>
                <th>Display Name</th>
                <th>Start Type</th>
            </tr>
"@
                            
                            foreach ($service in $servicesData.StoppedServices) {
                                $htmlReport += @"
            <tr>
                <td>$($service.Name)</td>
                <td>$($service.DisplayName)</td>
                <td>$($service.StartType)</td>
            </tr>
"@
                            }
                            
                            $htmlReport += @"
        </table>
"@
                        }
                        
                        $htmlReport += @"
    </div>
"@
                    }
                    
                    # Add inventory section if requested
                    if ($ReportContent -contains "Inventory" -and $reportData.Inventory[$computer]) {
                        $htmlReport += @"
    <div class="section">
        <h2>System Inventory</h2>
"@
                        
                        $diskData = $reportData.Inventory[$computer]["Disk"]
                        if ($diskData -and $diskData -isnot [string]) {
                            $htmlReport += @"
        <h3>Disk Information</h3>
        <table>
            <tr>
                <th>Drive</th>
                <th>Label</th>
                <th>Size (GB)</th>
                <th>Free Space (GB)</th>
                <th>Free (%)</th>
            </tr>
"@
                            
                            foreach ($volume in $diskData.Volumes) {
                                $htmlReport += @"
            <tr>
                <td>$($volume.DriveLetter)</td>
                <td>$($volume.FileSystemLabel)</td>
                <td>$($volume.SizeGB)</td>
                <td>$($volume.FreeSpaceGB)</td>
                <td>$($volume.FreePercent)%</td>
            </tr>
"@
                            }
                            
                            $htmlReport += @"
        </table>
"@
                        }
                        
                        $networkData = $reportData.Inventory[$computer]["Network"]
                        if ($networkData -and $networkData -isnot [string]) {
                            $htmlReport += @"
        <h3>Network Configuration</h3>
        <table>
            <tr>
                <th>Interface</th>
                <th>IP Address</th>
                <th>Gateway</th>
            </tr>
"@
                            
                            foreach ($interface in $networkData.IPConfiguration) {
                                $htmlReport += @"
            <tr>
                <td>$($interface.InterfaceAlias)</td>
                <td>$($interface.IPv4Address)</td>
                <td>$($interface.IPv4DefaultGateway.NextHop)</td>
            </tr>
"@
                            }
                            
                            $htmlReport += @"
        </table>
"@
                        }
                        
                        $softwareData = $reportData.Inventory[$computer]["Software"]
                        if ($softwareData -and $softwareData -isnot [string]) {
                            $htmlReport += @"
        <h3>Installed Software (Top 20 by Name)</h3>
        <table>
            <tr>
                <th>Name</th>
                <th>Version</th>
                <th>Publisher</th>
            </tr>
"@
                            
                            foreach ($software in ($softwareData.InstalledSoftware | Sort-Object DisplayName | Select-Object -First 20)) {
                                $htmlReport += @"
            <tr>
                <td>$($software.DisplayName)</td>
                <td>$($software.DisplayVersion)</td>
                <td>$($software.Publisher)</td>
            </tr>
"@
                            }
                            
                            $htmlReport += @"
        </table>
        
        <h3>Recently Installed Updates (Top 10)</h3>
        <table>
            <tr>
                <th>Hotfix ID</th>
                <th>Description</th>
                <th>Installed On</th>
            </tr>
"@
                            
                            foreach ($hotfix in ($softwareData.InstalledHotfixes | Select-Object -First 10)) {
                                $htmlReport += @"
            <tr>
                <td>$($hotfix.HotFixID)</td>
                <td>$($hotfix.Description)</td>
                <td>$($hotfix.InstalledOn)</td>
            </tr>
"@
                            }
                            
                            $htmlReport += @"
        </table>
"@
                        }
                        
                        $htmlReport += @"
    </div>
"@
                    }
                    
                    # Add HTML footer
                    $htmlReport += @"
</body>
</html>
"@
                    
                    # Save the HTML report
                    $htmlReport | Out-File -FilePath $reportFilePath -Encoding utf8
                }
                
                "CSV" {
                    Write-Verbose "Generating CSV reports for $computer..."
                    $baseReportPath = $reportFilePath
                    
                    # Handle each section as a separate CSV file
                    if ($ReportContent -contains "ServerInfo" -and $reportData.ServerInfo[$computer]) {
                        $sysInfoPath = "${baseReportPath}_SystemInfo.csv"
                        $sysInfo = $reportData.ServerInfo[$computer]["System"]
                        if ($sysInfo -and $sysInfo -isnot [string]) {
                            $sysInfo | Export-Csv -Path $sysInfoPath -NoTypeInformation
                        }
                        
                        $hwInfoPath = "${baseReportPath}_HardwareInfo.csv"
                        $hwInfo = $reportData.ServerInfo[$computer]["Hardware"]
                        if ($hwInfo -and $hwInfo -isnot [string]) {
                            $hwInfo | Export-Csv -Path $hwInfoPath -NoTypeInformation
                        }
                    }
                    
                    if ($ReportContent -contains "Performance" -and $reportData.Performance[$computer]) {
                        $perfData = $reportData.Performance[$computer]
                        if ($perfData -and $perfData -isnot [string]) {
                            foreach ($metricCategory in $perfData.Keys) {
                                $perfPath = "${baseReportPath}_${metricCategory}Metrics.csv"
                                $metricResults = @()
                                
                                foreach ($counterKey in $perfData[$metricCategory].Keys) {
                                    $counterSamples = $perfData[$metricCategory][$counterKey]
                                    $avgValue = ($counterSamples | Measure-Object -Property Value -Average).Average
                                    
                                    $metricResults += [PSCustomObject]@{
                                        Counter = $counterSamples[0].Counter
                                        Instance = $counterSamples[0].Instance
                                        AverageValue = [math]::Round($avgValue, 2)
                                    }
                                }
                                
                                $metricResults | Export-Csv -Path $perfPath -NoTypeInformation
                            }
                        }
                    }
                    
                    if ($ReportContent -contains "Services" -and $reportData.Services[$computer]) {
                        $servicesData = $reportData.Services[$computer]["Services"]
                        if ($servicesData -and $servicesData -isnot [string]) {
                            $runningServicesPath = "${baseReportPath}_RunningServices.csv"
                            $servicesData.RunningServices | Export-Csv -Path $runningServicesPath -NoTypeInformation
                            
                            $stoppedServicesPath = "${baseReportPath}_StoppedServices.csv"
                            $servicesData.StoppedServices | Export-Csv -Path $stoppedServicesPath -NoTypeInformation
                        }
                    }
                    
                    if ($ReportContent -contains "Inventory" -and $reportData.Inventory[$computer]) {
                        $diskData = $reportData.Inventory[$computer]["Disk"]
                        if ($diskData -and $diskData -isnot [string]) {
                            $volumesPath = "${baseReportPath}_Volumes.csv"
                            $diskData.Volumes | Export-Csv -Path $volumesPath -NoTypeInformation
                            
                            $physicalDisksPath = "${baseReportPath}_PhysicalDisks.csv"
                            $diskData.PhysicalDisks | Export-Csv -Path $physicalDisksPath -NoTypeInformation
                        }
                        
                        $networkData = $reportData.Inventory[$computer]["Network"]
                        if ($networkData -and $networkData -isnot [string]) {
                            $netConfigPath = "${baseReportPath}_NetworkConfig.csv"
                            $networkData.IPConfiguration | Export-Csv -Path $netConfigPath -NoTypeInformation
                        }
                        
                        $softwareData = $reportData.Inventory[$computer]["Software"]
                        if ($softwareData -and $softwareData -isnot [string]) {
                            $softwarePath = "${baseReportPath}_InstalledSoftware.csv"
                            $softwareData.InstalledSoftware | Export-Csv -Path $softwarePath -NoTypeInformation
                            
                            $hotfixPath = "${baseReportPath}_InstalledHotfixes.csv"
                            $softwareData.InstalledHotfixes | Export-Csv -Path $hotfixPath -NoTypeInformation
                        }
                    }
                }
                
                "XML" {
                    $reportFilePath += ".xml"
                    Write-Verbose "Generating XML report for $computer..."
                    
                    # Filter out only the data for this computer
                    $computerData = @{
                        ComputerName = $computer
                        ReportGenerated = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
                    }
                    
                    if ($ReportContent -contains "ServerInfo" -and $reportData.ServerInfo[$computer]) {
                        $computerData.ServerInfo = $reportData.ServerInfo[$computer]
                    }
                    
                    if ($ReportContent -contains "Performance" -and $reportData.Performance[$computer]) {
                        $computerData.Performance = $reportData.Performance[$computer]
                    }
                    
                    if ($ReportContent -contains "Services" -and $reportData.Services[$computer]) {
                        $computerData.Services = $reportData.Services[$computer]
                    }
                    
                    if ($ReportContent -contains "Inventory" -and $reportData.Inventory[$computer]) {
                        $computerData.Inventory = $reportData.Inventory[$computer]
                    }
                    
                    # Convert to XML and save
                    $computerData | Export-Clixml -Path $reportFilePath
                }
            }
            
            Write-Verbose "Report saved to $reportFilePath"
        }
    }
    
    end {
        Write-Verbose "Server reports generated successfully."
        return $true
    }
}