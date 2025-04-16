# MECM Force Update Script

## Overview

This PowerShell script automates the process of triggering Windows updates on multiple servers managed by Microsoft Endpoint Configuration Manager (MECM, formerly SCCM). It's designed to help IT administrators expedite the patching process during maintenance windows by forcing update evaluation and installation across multiple systems simultaneously.

## Features

- **Multi-threaded execution**: Process multiple servers in parallel to drastically reduce overall patching time
- **CSV-based targeting**: Target specific servers by listing them in a CSV file
- **Comprehensive logging**: Detailed logs showing operation status, timing, and results
- **Credential support**: Run with specific admin credentials as needed
- **Fallback mechanisms**: Try alternative methods if the primary approach fails
- **Progress tracking**: Real-time status updates as servers are processed
- **Resource management**: Control the number of concurrent operations to manage system load

## Requirements

- PowerShell 5.1 or higher
- MECM client installed on target servers
- Appropriate permissions to trigger MECM client actions
- Network connectivity to target servers

## Usage

### Basic Usage

```powershell
.\Force-MECMUpdates-MultiThreaded.ps1 -CsvPath "C:\Temp\servers.csv"
```
Note: The default logging location is C:\temp\MECMUpdateLogs\ <br>
(You may want to create the directy ahead of running the script to confirm the path)
### With Custom Parameters

```powershell
# Use specific credentials
$creds = Get-Credential
.\Force-MECMUpdates-MultiThreaded.ps1 -CsvPath "C:\Temp\servers.csv" -Credential $creds

# Change maximum concurrent threads (default is 10)
.\Force-MECMUpdates-MultiThreaded.ps1 -CsvPath "C:\Temp\servers.csv" -MaxThreads 20

# Specify a custom log directory
.\Force-MECMUpdates-MultiThreaded.ps1 -CsvPath "C:\Temp\servers.csv" -LogPath "D:\Logs\Patching"

# Combine multiple parameters
.\Force-MECMUpdates-MultiThreaded.ps1 -CsvPath "C:\Temp\servers.csv" -Credential $creds -MaxThreads 15 -LogPath "D:\Logs\Patching"
```

## CSV File Format

The script expects a CSV file with server names listed in a column named either `ServerName`, `ComputerName`, or `Name`. Example:

```csv
ServerName
server01.domain.com
server02.domain.com
server03.domain.com
```

## How It Works

The script performs the following actions for each server:

1. Verifies the server is online and reachable
2. Triggers the MECM Software Update Scan Cycle (Schedule ID: 00000000-0000-0000-0000-000000000113)
3. Waits for the scan to complete
4. Triggers the MECM Software Update Deployment Evaluation Cycle (Schedule ID: 00000000-0000-0000-0000-000000000108)
5. Checks for available updates using Microsoft.CCM.UpdatesStore
6. Initiates installation of pending updates
7. Records the result in the log file

Each of these operations is performed in parallel across multiple servers for maximum efficiency.

## Best Practices

- Test on a small group of servers before running against your entire environment
- Start with a low number of threads (5-10) and increase based on performance
- Run during an approved change window
- Monitor server status during and after patching
- Review logs for any servers that may need manual intervention

## Troubleshooting

If you encounter issues:

- Verify MECM client is installed and running on target servers
- Check network connectivity to the servers
- Ensure you have appropriate permissions
- Verify MECM client health on problem servers
- Check the detailed logs for specific error messages

## Notes

- This script triggers the update process but doesn't wait for completion
- Some updates may require a reboot to complete installation
- The script doesn't force reboots - servers will follow their configured MECM reboot policy
