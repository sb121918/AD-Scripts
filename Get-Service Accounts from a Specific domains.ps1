# Load GUI assembly
Add-Type -AssemblyName Microsoft.VisualBasic

# 1. Prompt for Domain
$TargetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the domain to search", "Service Account Search", "MTVN.ad.viacom.com")

if ([string]::IsNullOrWhiteSpace($TargetDomain)) { exit }

Write-Host "Searching ${TargetDomain} for accounts containing 'Service Account' or 'SVC'..." -ForegroundColor Cyan

try {
    # 2. LDAP Filter logic:
    # EA6 contains "Service Account" OR EmployeeID contains "SVC" OR EmployeeNumber contains "SVC"
    $LdapFilter = "(|(extensionAttribute6=*Service Account*)(employeeID=*SVC*)(employeeNumber=*SVC*))"

    # 3. Fetch users with the required attributes
    $Users = Get-ADUser -LDAPFilter $LdapFilter -Server $TargetDomain -Properties `
        extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, Enabled

    if ($null -eq $Users) {
        Write-Warning "No accounts found matching these 'Contains' criteria in ${TargetDomain}."
    } else {
        # 4. Process the data
        $Results = $Users | Select-Object `
            Name, 
            SamAccountName, 
            Enabled,
            @{Name="EA6_Value"; Expression={$_.extensionAttribute6}},
            @{Name="EmployeeID"; Expression={$_.employeeID}},
            @{Name="EmployeeNumber"; Expression={$_.employeeNumber}},
            Title, 
            Department,
            Description,
            @{Name="CreationDate"; Expression={$_.whenCreated}},
            DistinguishedName

        # 5. Show results
        Write-Host "Success: Found $($Users.Count) accounts." -ForegroundColor Green
        $Results | Out-GridView -Title "Service Account Search Results"

        # 6. Export to Desktop
        $Path = "$env:USERPROFILE\Desktop\Service_Account_Search_$(Get-Date -Format 'yyyyMMdd').csv"
        $Results | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to Desktop: $Path" -ForegroundColor Yellow
    }
}
catch {
    # FIXED: Using ${} to prevent Drive Provider error with the colon
    Write-Error "An error occurred connecting to ${TargetDomain}: $($_.Exception.Message)"
}