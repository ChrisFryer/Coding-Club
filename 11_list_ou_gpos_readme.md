# Active Directory GPO List Script

A PowerShell script to list all Group Policy Objects (GPOs) linked to a specific Organizational Unit (OU) in Active Directory, including their link order, status, enforcement settings, and scope.

## Features

- Lists all GPOs linked to a specified OU
- Displays GPO link order (processing order)
- Shows whether each GPO link is enabled or disabled
- Indicates if GPO links are enforced
- Shows the scope of each GPO (all OUs where it's linked)
- Exports results to a timestamped CSV file
- Color-coded console output for easy reading

## Prerequisites

- Windows PowerShell 5.1 or later
- Active Directory PowerShell module
- Group Policy PowerShell module
- Appropriate permissions to read AD and GPO information

### Installing Required Modules

The required modules are typically installed with Remote Server Administration Tools (RSAT):

**Windows 10/11:**
```powershell
# Install RSAT features
Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
Add-WindowsCapability -Online -Name Rsat.GroupPolicy.Management.Tools~~~~0.0.1.0
```

**Windows Server:**
```powershell
Install-WindowsFeature RSAT-AD-PowerShell, GPMC
```

## Usage

1. **Update the OU Path** in the script (line 10):
   ```powershell
   $OUPath = "OU=YourOU,OU=ParentOU,DC=domain,DC=com"
   ```

2. **Run the script** with appropriate permissions:
   ```powershell
   .\list_ou_gpos.ps1
   ```

3. **Review the output** in the console and check the generated CSV file

## How to Find Your OU Path

There are several methods to obtain the correct Distinguished Name (DN) for your OU:

### Method 1: Active Directory Users and Computers (ADUC)

1. Open **Active Directory Users and Computers**
2. Navigate to **View** → Enable **Advanced Features**
3. Right-click on your target OU → **Properties**
4. Go to the **Attribute Editor** tab
5. Find the **distinguishedName** attribute
6. Copy the value (e.g., `OU=TestOnly,OU=Users,OU=Department,DC=domain,DC=com`)

### Method 2: PowerShell

```powershell
# List all OUs in your domain
Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Sort-Object Name

# Search for a specific OU by name
Get-ADOrganizationalUnit -Filter "Name -eq 'YourOUName'" | Select-Object DistinguishedName

# Browse interactively (displays in a grid view)
Get-ADOrganizationalUnit -Filter * | Select-Object Name, DistinguishedName | Out-GridView
```

### Method 3: Using ADSI Edit

1. Open **ADSI Edit** (adsiedit.msc)
2. Connect to your domain
3. Navigate to your OU
4. Right-click → **Properties**
5. The **distinguishedName** is shown at the top

### Understanding the OU Path Format

The Distinguished Name follows this hierarchical format (read right to left):

```
OU=SubOU,OU=ParentOU,OU=DepartmentOU,DC=domain,DC=company,DC=com
```

- **OU=** Organizational Unit
- **DC=** Domain Component (parts of your domain name)
- Components are listed from most specific (left) to least specific (right)

**Example:** For `OU=TestOnly,OU=Users,OU=Department,DC=med,DC=fusetech,DC=com,DC=au`
- The most specific OU is "TestOnly"
- It's inside "Users"
- Which is inside "Department"
- In the domain "med.fusetech.com.au"

## Output Information

### Console Output

The script displays:
- OU name and canonical path
- Each linked GPO with its:
  - Link order number
  - Display name
  - Enabled/disabled status
  - Enforcement status
  - Scope (other OUs where it's linked)

### CSV Export

A timestamped CSV file is created with columns:
- Link Order
- GPO Name
- Enabled
- Enforced
- Scope (Linked OUs)

Filename format: `GPO_List_YYYYMMDD_HHMMSS.csv`

## Understanding GPO Settings

### Link Order
Lower numbers are processed first. GPOs with higher numbers (processed later) can override settings from earlier GPOs.

### Enabled
- **Yes**: The GPO link is active and will be applied
- **No**: The GPO link is disabled and will not be applied

### Enforced
- **Yes**: This GPO cannot be blocked by child OUs
- **No**: Child OUs can block this GPO

### Scope
Shows all OUs where this GPO is linked, helping you understand its full impact across your domain.

## Troubleshooting

### Common Errors

**"Cannot find object"**
- Verify the OU path is correct
- Ensure you have read permissions for the OU

**"Module not found"**
- Install RSAT tools (see Prerequisites)
- Run PowerShell as Administrator

**"Access Denied"**
- You need read permissions for AD and GPO
- Contact your domain administrator

### Execution Policy

If you cannot run the script, you may need to adjust your execution policy:

```powershell
# View current policy
Get-ExecutionPolicy

# Allow local scripts (run as Administrator)
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Permissions Required

The account running this script needs:
- Read access to Active Directory
- Read access to Group Policy Objects
- Typically requires Domain User membership at minimum
- Some environments may require additional permissions

## License

This script is provided as-is for use in Active Directory environments.

## Contributing

Feel free to submit issues or pull requests for improvements.
