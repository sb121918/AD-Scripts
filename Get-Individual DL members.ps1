# --- Configuration ---
Add-Type -AssemblyName Microsoft.VisualBasic
$ExportToCSV = $true
$MemberExportPath = "$HOME\Desktop\DL_Member_Details.csv"

# --- Popup Input Window ---
$Title = "Distribution List Lookup"
$Msg   = "Enter DL Name, Email ID, or use wildcards (e.g. *Finance*):"
$Query = [Microsoft.VisualBasic.Interaction]::InputBox($Msg, $Title)

if ([string]::IsNullOrWhiteSpace($Query)) {
    Write-Host "Search cancelled or empty input." -ForegroundColor Yellow
    exit
}

Write-Host "`nSearching Active Directory for: '$Query'..." -ForegroundColor Cyan

# --- Step 1: Find the Groups ---
$Filter = "Name -like '$Query' -or DisplayName -like '$Query' -or mail -like '$Query'"

try {
    $Groups = Get-ADGroup -Filter $Filter -Properties mail, DisplayName, GroupCategory | 
              Where-Object { $_.GroupCategory -eq "Distribution" } |
              Select-Object DisplayName, Name, @{Name="EmailAddress"; Expression={$_.mail}}, DistinguishedName

    if (-not $Groups) {
        Write-Host "No results found for '$Query'." -ForegroundColor Red
        exit
    }

    # --- Step 2: User Selection ---
    $SelectedGroup = $Groups | Out-GridView -Title "Select ONE Distribution List and hit OK" -OutputMode Single

    if (-not $SelectedGroup) {
        Write-Host "No group selected." -ForegroundColor Yellow
        exit
    }

    # --- Step 3: Get Members (Robust Logic) ---
    Write-Host "Fetching members for: $($SelectedGroup.DisplayName)..." -ForegroundColor Cyan
    
    $GroupObj = Get-ADGroup -Identity $SelectedGroup.DistinguishedName -Properties member
    $MemberList = New-Object System.Collections.Generic.List[PSObject]

    foreach ($MemberDN in $GroupObj.member) {
        try {
            # Get-ADObject handles cross-domain and foreign security principals safely
            $Object = Get-ADObject -Identity $MemberDN -Properties Name, Title, Department, mail, objectClass
            
            $MemberData = [PSCustomObject]@{
                Name       = $Object.Name
                Type       = $Object.objectClass
                Title      = if ($Object.Title) { $Object.Title } else { "N/A" }
                Department = if ($Object.Department) { $Object.Department } else { "N/A" }
                Email      = if ($Object.mail) { $Object.mail } else { "N/A" }
                DistinguishedName = $Object.DistinguishedName
            }
            $MemberList.Add($MemberData)
        } catch {
            # This handles deleted objects or unreachable forest members
            $MemberList.Add([PSCustomObject]@{
                Name       = "Unresolved/Foreign Member"
                Type       = "Unknown"
                Title      = "N/A"
                Department = "N/A"
                Email      = $MemberDN
                DistinguishedName = "Error Resolving"
            })
        }
    }

    # --- Step 4: Output & Export ---
    if ($MemberList.Count -gt 0) {
        Write-Host "Success! Found $($MemberList.Count) members." -ForegroundColor Green
        
        $MemberList | Out-GridView -Title "Members of $($SelectedGroup.DisplayName)"

        if ($ExportToCSV) {
            $MemberList | Export-Csv -Path $MemberExportPath -NoTypeInformation -Encoding utf8
            Write-Host "`nMember list exported to: $MemberExportPath" -ForegroundColor Green
        }
    } else {
        Write-Host "The selected group is empty." -ForegroundColor Yellow
    }

} catch {
    Write-Host "Critical Error: $($_.Exception.Message)" -ForegroundColor Red
}