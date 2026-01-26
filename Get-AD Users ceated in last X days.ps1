# Load AD Module
Import-Module ActiveDirectory

# UI for Input
Add-Type -AssemblyName Microsoft.VisualBasic
$days = [Microsoft.VisualBasic.Interaction]::InputBox("Enter number of days to look back:", "AD Forest Report", "7")

if ([string]::IsNullOrWhiteSpace($days)) { Write-Host "Cancelled."; exit }

$dateThreshold = (Get-Date).AddDays(-$days)
$exportPath = "$env:USERPROFILE\Desktop\AD_NewUsers_Report_$((Get-Date).ToString('yyyyMMdd')).csv"
$results = @()

# Get all domains in the forest
$allDomains = (Get-ADForest).Domains

Write-Host "Starting forest scan for users created since $dateThreshold..." -ForegroundColor Cyan

foreach ($domain in $allDomains) {
    Write-Host "Scanning Domain: $domain" -ForegroundColor Gray
    try {
        # Define all requested properties
        $props = @(
            "GivenName", "Surname", "DisplayName", "Title", "UserPrincipalName", 
            "SamAccountName", "EmailAddress", "Department", "EmployeeID", 
            "EmployeeNumber", "EmployeeType", "Description", "Manager", 
            "whenCreated", "PasswordNeverExpires", "PasswordExpired", 
            "msExchExtensionAttribute27", "msExchExtensionAttribute28", 
            "extensionAttribute5", "extensionAttribute6", "DistinguishedName"
        )

        $users = Get-ADUser -Filter 'whenCreated -ge $dateThreshold' -Server $domain -Properties $props | ForEach-Object {
            
            # Resolve Manager Name
            $mgrName = "N/A"
            if ($_.Manager) {
                try { $mgrName = (Get-ADUser -Identity $_.Manager -Server $domain).DisplayName } catch { $mgrName = "Unknown/Cross-Domain" }
            }

            # Map to Spreadsheet Columns
            [PSCustomObject]@{
                "First Name"                    = $_.GivenName
                "Last Name"                     = $_.Surname
                "Display Name"                  = $_.DisplayName
                "Employee Name"                 = $_.DisplayName
                "Title"                         = $_.Title
                "User Logon Name (UPN)"         = $_.UserPrincipalName
                "SamAccountName"                = $_.SamAccountName
                "Email Address"                 = $_.EmailAddress
                "Department"                    = $_.Department
                "Employee ID"                   = $_.EmployeeID
                "Employee Number"               = $_.EmployeeNumber
                "Employee Type"                 = $_.EmployeeType
                "Description"                   = $_.Description
                "Manager Name"                  = $mgrName
                "Date Created"                  = $_.whenCreated
                "Password Never Expire"         = $_.PasswordNeverExpires
                "Reset Password at First Logon" = $_.PasswordExpired
                "MsExchExtensionAttribute27"    = $_.msExchExtensionAttribute27
                "MsExchExtensionAttribute28"    = $_.msExchExtensionAttribute28
                "Custom Attribute 5"            = $_.extensionAttribute5
                "Custom Attribute 6"            = $_.extensionAttribute6
                "Distinguished Name"            = $_.DistinguishedName
                "Domain"                        = $domain
            }
        }
        $results += $users
    } catch {
        Write-Warning "Failed to query $domain. Ensure you have connectivity and permissions."
    }
}

# Export to CSV and Open
if ($results) {
    $results | Export-Csv -Path $exportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Success! Report saved to Desktop: $exportPath" -ForegroundColor Green
    Invoke-Item $exportPath
} else {
    Write-Host "No users found for the selected timeframe." -ForegroundColor Yellow
}