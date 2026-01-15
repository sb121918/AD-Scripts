# Load GUI assembly
Add-Type -AssemblyName Microsoft.VisualBasic

# 1. Get Forest Information
try {
    $Forest = Get-ADForest
    $ForestRoot = $Forest.RootDomain
}
catch {
    Write-Error "Unable to connect to Active Directory. Ensure you have appropriate network access."
    exit
}

Write-Host "Connecting to Global Catalog on ${ForestRoot} (Port 3268)..." -ForegroundColor Cyan

try {
    # 2. Define the Search Criteria (EA6 or SVC tag)
    $LdapFilter = "(|(extensionAttribute6=*Non-Human Service Account*)(employeeID=*SVC*)(employeeNumber=*SVC*))"

    # 3. Query the Global Catalog (Port 3268)
    $Users = Get-ADUser -LDAPFilter $LdapFilter -Server "${ForestRoot}:3268" -Properties `
        extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, whenChanged, Enabled, Manager

    # 4. Filter for accounts where Manager is null or empty
    $OrphanedAccounts = $Users | Where-Object { [string]::IsNullOrWhiteSpace($_.Manager) }

    if ($null -eq $OrphanedAccounts -or ($OrphanedAccounts | Measure-Object).Count -eq 0) {
        Write-Warning "No existing orphaned service accounts found in the forest."
    } else {
        # 5. Process results
        $Results = $OrphanedAccounts | ForEach-Object {
            
            # Clean Domain Parsing
            $DnParts = $_.DistinguishedName.Split(',') | Where-Object { $_ -like "DC=*" }
            $CleanDomain = ($DnParts -replace "DC=", "") -join "."

            [PSCustomObject]@{
                Name              = $_.Name
                SamAccountName    = $_.SamAccountName
                AccountType       = "Service Account (Orphaned)"
                Enabled           = $_.Enabled
                Domain            = $CleanDomain
                Created           = $_.whenCreated
                LastModified      = $_.whenChanged
                EA6_Value         = $_.extensionAttribute6
                EmployeeID        = $_.employeeID
                EmployeeNumber    = $_.employeeNumber
                ManagerStatus     = "NOT ASSIGNED"
                Title             = $_.Title
                Department        = $_.Department
                DistinguishedName = $_.DistinguishedName
            }
        }

        # 6. Output Results
        Write-Host "Success! Found $($Results.Count) orphaned service accounts." -ForegroundColor Green
        
        # Display in GridView
        $Results | Out-GridView -Title "Full Orphaned Service Account Audit - ${ForestRoot}"
        
        # Export to Desktop
        $Path = "$env:USERPROFILE\Desktop\Full_Orphaned_SvcAccounts_$(Get-Date -Format 'yyyyMMdd').csv"
        $Results | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to: $Path" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Error querying Global Catalog: $($_.Exception.Message)"
}