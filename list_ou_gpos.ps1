# Script to list all Group Policy Objects linked to a specific OU
# This will show the actual GPO Display Names

# Import required modules
Import-Module ActiveDirectory
Import-Module GroupPolicy

# Specify your OU path - UPDATE THIS!
# Example: "OU=TestOnly,OU=Users,OU=Department,DC=domain,DC=com"
$OUPath = "OU=TestOnly,OU=Users,OU=Department,DC=med,DC=fusetech,DC=com,DC=au"

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Group Policies Linked to OU" -ForegroundColor Cyan
Write-Host "========================================`n" -ForegroundColor Cyan

try {
    # Get the OU and its linked GPOs
    $OU = Get-ADOrganizationalUnit -Identity $OUPath -Properties gpLink, CanonicalName
    
    Write-Host "OU Name: $($OU.Name)" -ForegroundColor Green
    Write-Host "Path: $($OU.CanonicalName)" -ForegroundColor Gray
    Write-Host ""
    
    if ([string]::IsNullOrEmpty($OU.gpLink)) {
        Write-Host "No Group Policies are linked to this OU.`n" -ForegroundColor Yellow
        exit
    }
    
    # Parse the gpLink attribute to extract GPO GUIDs
    $gpLinks = @()
    $linkOrder = 1
    
    # Split and process each GPO link
    $OU.gpLink -split "\]\[" | ForEach-Object {
        $link = $_ -replace "[\[\]]", ""
        
        if ($link -match "cn=({[^}]+})") {
            $guid = $matches[1]
            
            # Determine link properties
            $linkEnabled = if ($link -match ";1$|;3$") { "No" } else { "Yes" }
            $enforced = if ($link -match ";2$|;3$") { "Yes" } else { "No" }
            
            $gpLinks += [PSCustomObject]@{
                Order = $linkOrder
                GUID = $guid
                Enabled = $linkEnabled
                Enforced = $enforced
            }
            
            $linkOrder++
        }
    }
    
    # Get GPO details for each linked policy
    Write-Host "Linked Group Policies:" -ForegroundColor Yellow
    Write-Host ("=" * 80) -ForegroundColor Gray
    
    $results = @()
    
    foreach ($link in $gpLinks) {
        try {
            $gpo = Get-GPO -Guid $link.GUID -ErrorAction Stop
            
            # Get all OUs where this GPO is linked (Scope)
            $gpoReport = [xml](Get-GPOReport -Guid $link.GUID -ReportType Xml)
            $linkedOUs = $gpoReport.GPO.LinksTo | ForEach-Object { 
                $somPath = $_.SOMPath
                # Extract just the last OU name from the path
                # Split by / and get the last element
                $pathParts = $somPath -split '/'
                $pathParts[-1]
            }
            $scopeList = if ($linkedOUs) { ($linkedOUs | Select-Object -Unique) -join "; " } else { "None" }
            
            $result = [PSCustomObject]@{
                'Link Order' = $link.Order
                'GPO Name' = $gpo.DisplayName
                'Enabled' = $link.Enabled
                'Enforced' = $link.Enforced
                'Scope (Linked OUs)' = $scopeList
            }
            
            $results += $result
            
            # Display each GPO
            Write-Host "`n[$($link.Order)] $($gpo.DisplayName)" -ForegroundColor White
            Write-Host "    Enabled: $($link.Enabled)  |  Enforced: $($link.Enforced)" -ForegroundColor Gray
            Write-Host "    Scope: $scopeList" -ForegroundColor DarkGray
            
        }
        catch {
            Write-Warning "Could not retrieve GPO with GUID: $($link.GUID)"
        }
    }
    
    Write-Host "`n" 
    Write-Host ("=" * 80) -ForegroundColor Gray
    Write-Host "Total GPOs linked: $($results.Count)" -ForegroundColor Green
    
    # Export to CSV with proper formatting
    if ($results.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $exportPath = ".\GPO_List_$timestamp.csv"
        $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $exportPath" -ForegroundColor Cyan
    }
    
}
catch {
    Write-Error "Error: $_"
    Write-Host "`nMake sure:" -ForegroundColor Yellow
    Write-Host "  1. The OU path is correct" -ForegroundColor Yellow
    Write-Host "  2. You have permissions to read AD and GPO information" -ForegroundColor Yellow
    Write-Host "  3. The Active Directory and GroupPolicy modules are installed`n" -ForegroundColor Yellow
}

Write-Host ""
