# Load GUI assembly
Add-Type -AssemblyName Microsoft.VisualBasic

# 1. Prompt for Domain
$TargetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the domain to search", "Identity Match Search", "MTVN.ad.viacom.com")

if ([string]::IsNullOrWhiteSpace($TargetDomain)) { exit }

Write-Host "Searching ${TargetDomain} for accounts where EA6 equals 'Human Primary Identity SF Match'..." -ForegroundColor Cyan

try {
    # 2. LDAP Filter logic:
    # Exact match for extensionAttribute6. Note: No wildcards (*) used here.
    $LdapFilter = "(extensionAttribute6=Human Primary Identity SF Match)"

    # 3. Fetch users with the required attributes
    $Users = Get-ADUser -LDAPFilter $LdapFilter -Server $TargetDomain -Properties `
        extensionAttribute6, employeeID, employeeNumber, Title, Department, Description, whenCreated, Enabled

    if ($null -eq $Users -or $Users.Count -eq 0) {
        Write-Warning "No accounts found with an exact EA6 match in ${TargetDomain}."
    } else {
        # 4. Process the data for output
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

        # 5. Show results in GridView
        Write-Host "Success: Found $($Users.Count) matching identity accounts." -ForegroundColor Green
        $Results | Out-GridView -Title "Identity Match Results: Human Primary Identity SF Match"

        # 6. Export to Desktop with specific filename
        $Path = "$env:USERPROFILE\Desktop\Identity_SF_Match_$(Get-Date -Format 'yyyyMMdd').csv"
        $Results | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to Desktop: $Path" -ForegroundColor Yellow
    }
}
catch {
    # Using ${} to prevent Drive Provider error with the colon
    Write-Error "An error occurred connecting to ${TargetDomain}: $($_.Exception.Message)"
}