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

# 2. Prompt for Timeframe (Optional check for newly created accounts)
$DaysInput = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the number of days to look back for accounts (e.g., 30, 365, or 9999 for all time):", "Creation Timeframe", "365")

$DaysCount = 0 
if (-not [int]::TryParse($DaysInput, [ref]$DaysCount)) { $DaysCount = 365 }

# Calculate LDAP Date
$ThresholdDate = (Get-Date).AddDays(-$DaysCount)
$LdapDate = $ThresholdDate.ToString("yyyyMMddHHmmss.0Z")

Write-Host "Connecting to Global Catalog on ${ForestRoot}..." -ForegroundColor Cyan

try {
    # 3. Define the Search Criteria
    # EA6 = "Non-Human Service Account" OR ID contains "SVC" OR Number contains "SVC"
    $LdapFilter = "(&(whenCreated>=$LdapDate)(|(extensionAttribute6=*Non-Human Service Account*)(employeeID=*SVC*)(employeeNumber=*SVC*)))"

    # 4. Query the Global Catalog (Port 3268)
    # We must include 'Manager' in the properties list to check it
    $Users = Get-ADUser -LDAPFilter $LdapFilter -Server "${ForestRoot}:3268" -Properties `
        extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, Enabled, Manager

    # 5. Filter for accounts where Manager is NOT assigned
    $OrphanedAccounts = $Users | Where-Object { [string]::IsNullOrEmpty($_.Manager) }

    if ($null -eq $OrphanedAccounts -or ($OrphanedAccounts | Measure-Object).Count -eq 0) {
        Write-Warning "No service accounts found missing a manager within the last $DaysCount days."
    } else {
        # 6. Process results
        $Results = $OrphanedAccounts | ForEach-Object {
            
            # Clean Domain Parsing
            $DnParts = $_.DistinguishedName.Split(',') | Where-Object { $_ -like "DC=*" }
            $CleanDomain = ($DnParts -replace "DC=", "") -join "."

            [PSCustomObject]@{
                Name            = $_.Name
                SamAccountName  = $_.SamAccountName
                AccountType     = "Service Account (Unmanaged)"
                Enabled         = $_.Enabled
                Domain          = $CleanDomain
                Created         = $_.whenCreated
                EA6_Value       = $_.extensionAttribute6
                EmployeeID      = $_.employeeID
                EmployeeNumber  = $_.employeeNumber
                ManagerStatus   = "MISSING"
                Title           = $_.Title
                Department      = $_.Department
                DistinguishedName = $_.DistinguishedName
            }
        }

        # 7. Output Results
        Write-Host "Success! Found $($Results.Count) service accounts without a Manager." -ForegroundColor Green
        
        $Results | Out-GridView -Title "Service Accounts MISSING Manager: ${ForestRoot}"
        
        $Path = "$env:USERPROFILE\Desktop\Orphaned_Service_Accounts_$(Get-Date -Format 'yyyyMMdd').csv"
        $Results | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to Desktop: $Path" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Error querying Global Catalog: $($_.Exception.Message)"
}