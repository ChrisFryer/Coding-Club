# Windows Server Update Checker

A PowerShell GUI application for checking and reporting on Windows updates across multiple servers in your environment.

![Application Screenshot](screenshots/app_screenshot.png)

## Overview

The Windows Server Update Checker provides IT administrators with a simple interface to:

- Import server lists from CSV files
- Query update information from multiple servers simultaneously
- Filter and view installed updates by type
- Generate patch status reports with visualizations
- Export results to CSV for further analysis

## Version History

### v3.0 - Latest Release
- Added detailed reporting capabilities with visual charts
- Server patch status distribution (pie chart)
- Monthly patch trend analysis (bar chart)
- Status summaries with counts and classifications
- Report export functionality 
- Performance improvements

### v2.0
- Added ability to filter updates by type (Security Update, Hotfix, etc.)
- Added option to show only the latest update per server
- Returns all installed updates, not just those with "Update" description
- Improved UI layout and responsiveness
- Added "Apply Filters" and "Show All Results" buttons

### v1.0 - Initial Release
- Basic server update checking functionality
- CSV import of server lists
- Connection options for SSL and credentials
- Results display in a data grid
- Export to CSV capability

## Requirements

- Windows PowerShell 5.1 or higher
- Windows operating system (tested on Windows 10/11 and Windows Server 2016/2019/2022)
- Remote PowerShell access to target servers
- Appropriate permissions to query update information

## Installation

No installation required. Simply download the script file for the version you want to use and run it with PowerShell:

```powershell
.\WinUpdateAppv3.0.ps1
```

## Usage Instructions

### Importing Servers

1. Click "Import CSV" to select a CSV file containing server names
   - The application supports various column headers: "ComputerName", "ServerName", "Server", "Name", or "Computer"
   - If none of these column names are found, it will use the first column

### Connection Options

- **Use SSL**: Enable this for secure connections to remote servers
- **Use Custom Credentials**: Enable to specify alternate credentials for remote server access

### Querying Updates

1. Select one or more servers from the imported list (use Ctrl or Shift for multiple selection)
2. Click "Get Last Updates" to retrieve update information
3. View results in the grid on the right side

### Filtering Results (v2.0 and above)

1. Use the "Filter by Update Type" options to select specific update types
2. Enable "Latest Update Only" to see only the most recent update per server
3. Click "Apply Filters" to update the view based on your selections
4. Use "Show All Results" to revert to the unfiltered view

### Reporting (v3.0 only)

1. After retrieving update information, click "Produce Report"
2. The report window will display:
   - Summary statistics of server patch status
   - Patch status distribution chart
   - Monthly patch trend analysis
3. Click "Export Report" to save the report data as a CSV file

### Exporting Results

1. Click "Export Results" to save the query results as a CSV file
2. Choose a location and filename for the export

## Remote Server Requirements

- PowerShell Remoting must be enabled on target servers
- WinRM service must be running
- Appropriate firewall rules must allow PowerShell Remoting traffic
- The account running the script must have permissions to query update information via Get-HotFix

## Troubleshooting

### Common Issues

- **Connection Failures**: Ensure PowerShell Remoting is enabled on target servers and network connectivity exists
- **Access Denied**: Verify the account has appropriate permissions on the target servers
- **No Updates Displayed**: Confirm target servers have Windows updates installed and are accessible

### Error Messages in Results

When connection or permission issues occur, the application will display error information in the Status column of the results grid.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions to improve the Windows Server Update Checker are welcome. Please feel free to submit issues or pull requests to the repository.

## Acknowledgments

- Thanks to the PowerShell community for guidance and inspiration
- Icons and visual elements based on Windows Forms standard controls
