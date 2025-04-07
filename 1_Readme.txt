To install:

1. Copy the code to a working directory
2. Import-Module .\RemoteServerManagement.psm1
3. Start-RemoteServerManagement

Or use any of the individual functions directly:
# Test connection to remote servers
Test-RemoteServerConnection -ComputerName "server01", "server02"

# Get detailed system information
Get-RemoteServerInfo -ComputerName "server01" -InfoType System,Hardware