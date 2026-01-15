# Load GUI assembly for the prompt
Add-Type -AssemblyName Microsoft.VisualBasic

# 1. Prompt for Domain (Pre-filled with your specific domain)
$TargetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the domain to search", "Domain Search", "MTVN.ad.viacom.com")

if ([string]::IsNullOrWhiteSpace($TargetDomain)) { exit }

Write-Host "Connecting to $TargetDomain and collecting employee types..." -ForegroundColor Cyan

try {
    # 2. Get users and the attributes usually used for 'Type' (EmployeeType and EA6)
    $Users = Get-ADUser -Filter 'Enabled -eq $true' -Server $TargetDomain -Properties EmployeeType, extensionAttribute6, Title, Department

    # 3. Process results
    $Results = $Users | Select-Object `
        Name, 
        SamAccountName, 
        @{Name="EmployeeType_Field"; Expression={$_.EmployeeType}},
        @{Name="EA6_Value"; Expression={$_.extensionAttribute6}},
        Title, 
        Department,
        DistinguishedName

    # 4. Show the detailed list
    $Results | Out-GridView -Title "Employee Classifications in $TargetDomain"

    # 5. Summary View: Count how many of each type exist
    Write-Host "`n--- Summary of EmployeeType Field ---" -ForegroundColor Yellow
    $Results | Group-Object EmployeeType_Field -NoElement | Sort-Object Count -Descending | Format-Table Name, Count -AutoSize

    Write-Host "--- Summary of extensionAttribute6 (EA6) ---" -ForegroundColor Yellow
    $Results | Group-Object EA6_Value -NoElement | Sort-Object Count -Descending | Format-Table Name, Count -AutoSize

}
catch {
    Write-Error "Failed to connect to $TargetDomain. Verify VPN/Network and RSAT tools."
}