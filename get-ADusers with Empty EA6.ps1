# --- Configuration ---
$ExportToCSV = $true  
$FilePath = "$HOME\Desktop\Forest_Empty_EA6_Report.csv"

# --- Logic ---
Write-Host "Identifying all domains in the forest..." -ForegroundColor Cyan
try {
    $AllDomains = (Get-ADForest).Domains
} catch {
    Write-Host "Error: Could not retrieve forest information. Ensure you have RSAT installed." -ForegroundColor Red
    return
}

$MasterUserList = New-Object System.Collections.Generic.List[PSObject]

$ADProps = @(
    "extensionAttribute6", 
    "Manager", 
    "EmailAddress", 
    "PasswordLastSet", 
    "WhenCreated"
)

foreach ($Domain in $AllDomains) {
    Write-Host "Searching Domain: $Domain ..." -ForegroundColor White
    
    try {
        # Search each domain specifically
        $DomainUsers = Get-ADUser -Filter 'extensionAttribute6 -notlike "*"' -Server $Domain -Properties $ADProps | 
            Select-Object Name, 
                SamAccountName, 
                @{Name="EmailID"; Expression={$_.EmailAddress}},
                @{Name="Manager"; Expression={if($_.Manager){(Get-ADUser $_.Manager -Server $Domain).Name} else {"N/A"}}},
                @{Name="OU"; Expression={($_.DistinguishedName -split ',', 2)[1]}},
                @{Name="Domain"; Expression={$Domain}},
                @{Name="PasswordLastSet"; Expression={$_.PasswordLastSet}},
                @{Name="WhenCreated"; Expression={$_.WhenCreated}},
                extensionAttribute6

        if ($DomainUsers) {
            $MasterUserList.AddRange($DomainUsers)
            Write-Host "Found $($DomainUsers.Count) users in $Domain." -ForegroundColor Green
        }
    } catch {
        Write-Host "Warning: Could not contact or search domain $Domain. Skipping..." -ForegroundColor Yellow
    }
}

# --- Output ---
if ($MasterUserList.Count -gt 0) {
    # Display in console
    $MasterUserList | Out-GridView -Title "Forest-Wide Users with Empty EA6"
    $MasterUserList | Format-Table Name, EmailID, Domain, PasswordLastSet -AutoSize

    # Handle Export
    if ($ExportToCSV) {
        try {
            $MasterUserList | Export-Csv -Path $FilePath -NoTypeInformation -Encoding utf8
            Write-Host "`nSuccess! Full forest report exported to: $FilePath" -ForegroundColor Green
        } catch {
            Write-Host "`nError: Could not save file. Check if it's open in Excel." -ForegroundColor Red
        }
    }
} else {
    Write-Host "No users found with an empty extensionAttribute6 across the entire forest." -ForegroundColor Yellow
}