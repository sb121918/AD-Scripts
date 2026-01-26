# 1. Load the required assembly
try {
    Add-Type -AssemblyName Microsoft.VisualBasic
} catch {
    Write-Warning "Visual Basic assembly not found."
}

# 2. Prompt for Search Term
$SearchValue = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Name, Email, SAM, or Employee ID", "Forest-Wide Identity Search")

if ([string]::IsNullOrWhiteSpace($SearchValue)) { 
    Write-Host "No input detected. Exiting..." -ForegroundColor Yellow
    exit 
}

$WildcardSearch = "*$($SearchValue.Trim('*'))*"
Write-Host "Searching ALL Domains for: $WildcardSearch" -ForegroundColor Cyan

try {
    # 3. Get the Global Catalog Server
    $CurrentForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $GlobalCatalog = ($CurrentForest.GlobalCatalogs[0]).Name
    
    # 4. Filter and Search
    $Filter = "Name -like '$WildcardSearch' -or DisplayName -like '$WildcardSearch' -or mail -like '$WildcardSearch' -or SamAccountName -like '$WildcardSearch' -or employeeID -like '$WildcardSearch'"
    
    $ADUsers = Get-ADUser -Filter $Filter -Server "$($GlobalCatalog):3268" -Properties `
        DisplayName, EmailAddress, employeeID, employeeNumber, Title, Department, `
        proxyAddresses, MemberOf, Enabled

    if ($null -eq $ADUsers) {
        Write-Warning "No user found matching '$SearchValue'."
    } else {
        # 5. Build Initial List for Selection
        $InitialList = foreach ($User in $ADUsers) {
            $DomainDN = ($User.DistinguishedName -split 'DC=')[-2..-1] -join '.'
            [PSCustomObject]@{
                DisplayName    = $User.DisplayName
                SAM            = $User.SamAccountName
                Domain         = $DomainDN.Replace(',','')
                Email          = $User.EmailAddress
                Department     = $User.Department
                Status         = if ($User.Enabled) { "Enabled" } else { "Disabled" }
                # Keep the original object hidden for later
                _OriginalData  = $User
            }
        }

        # 6. CHOICE BOX: User selects the specific account
        $SelectedUser = $InitialList | Out-GridView -Title "SELECT A USER AND CLICK OK" -PassThru

        if ($null -ne $SelectedUser) {
            Write-Host "Retrieving full details for: $($SelectedUser.DisplayName)..." -ForegroundColor Green
            
            # 7. Process Final Detailed View
            $FullUser = $SelectedUser._OriginalData
            $Groups = $FullUser.MemberOf | ForEach-Object { ($_ -split ',')[0].Replace("CN=","") } | Sort-Object

            $FinalReport = [PSCustomObject]@{
                DisplayName       = $FullUser.DisplayName
                AccountStatus     = $SelectedUser.Status
                Domain            = $SelectedUser.Domain
                Username_SAM      = $FullUser.SamAccountName
                Email             = $FullUser.EmailAddress
                Title             = $FullUser.Title
                Department        = $FullUser.Department
                EmployeeID        = $FullUser.employeeID
                GroupCount        = $Groups.Count
                Groups            = $Groups -join " | "
                ProxyAddresses    = ($FullUser.proxyAddresses | Where-Object {$_ -like "smtp:*"}) -join "; "
                DistinguishedName = $FullUser.DistinguishedName
            }

            # Show the single selected user in a final grid
            $FinalReport | Out-GridView -Title "Final Details: $($FullUser.DisplayName)"
        } else {
            Write-Host "No user selected. Operation cancelled." -ForegroundColor Yellow
        }
    }
}
catch {
    Write-Error "Error: $($_.Exception.Message)"
}