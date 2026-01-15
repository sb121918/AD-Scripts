# Load GUI assembly
Add-Type -AssemblyName Microsoft.VisualBasic

# 1. Prompt for the Target Domain
$TargetDomain = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the FQDN of the domain you want to search (e.g. child.domain.com)", "Target Domain", "")

if ([string]::IsNullOrWhiteSpace($TargetDomain)) {
    Write-Warning "No domain entered. Exiting."
    exit
}

try {
    Write-Host "Connecting to $TargetDomain..." -ForegroundColor Cyan
    
    # 2. Query only Enabled users from the specific server/domain
    # We use -Properties * to ensure we get everything you might need
    $Users = Get-ADUser -Filter 'Enabled -eq $true' -Server $TargetDomain -Properties `
        DisplayName, EmailAddress, Title, Department, Company, LastLogonDate, whenCreated

    if ($null -eq $Users) {
        Write-Warning "No enabled users found in $TargetDomain."
    } else {
        # 3. Format the results
        $Results = $Users | Select-Object `
            @{Name="Domain"; Expression={$TargetDomain}},
            Name, 
            DisplayName, 
            SamAccountName, 
            UserPrincipalName, 
            EmailAddress, 
            Title, 
            Department, 
            Company,
            LastLogonDate,
            whenCreated

        # 4. Show in GridView
        Write-Host "Found $($Users.Count) enabled users." -ForegroundColor Green
        $Results | Out-GridView -Title "Enabled Users in $TargetDomain"

        # 5. Export to Desktop
        $FilePath = "$env:USERPROFILE\Desktop\Enabled_Users_$($TargetDomain).csv"
        $Results | Export-Csv -Path $FilePath -NoTypeInformation -Encoding UTF8
        Write-Host "Export saved to: $FilePath" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Could not connect to domain '$TargetDomain'. Ensure the name is correct and you have network connectivity."
}