# Load GUI assembly
Add-Type -AssemblyName Microsoft.VisualBasic

Write-Host "Starting Forest-wide search for all enabled accounts..." -ForegroundColor Cyan

try {
    # 1. Get all domains in the current Forest
    $ForestDomains = (Get-ADForest).Domains
    $MasterResults = @()

    foreach ($Domain in $ForestDomains) {
        Write-Host "--- Searching Domain: $Domain ---" -ForegroundColor Gray
        
        try {
            # 2. Fetch ENABLED users with all requested properties
            $Users = Get-ADUser -Filter 'Enabled -eq $true' -Server $Domain -Properties `
                GivenName, Surname, DisplayName, Title, UserPrincipalName, SamAccountName, `
                EmailAddress, Department, EmployeeID, EmployeeNumber, EmployeeType, `
                Description, Manager, whenCreated, PasswordNeverExpires, PasswordExpired, `
                msExchExtensionAttribute27, msExchExtensionAttribute28, extensionAttribute5, `
                extensionAttribute6, DistinguishedName

            if ($null -ne $Users) {
                Write-Host "Processing $($Users.Count) accounts..." -ForegroundColor White
                
                # 3. Process the data with custom mapping
                $CurrentDomainResults = foreach ($user in $Users) {
                    # Resolve Manager DN to a readable Name
                    $mgrName = $null
                    if ($user.Manager) {
                        try {
                            $mgrName = (Get-ADUser -Identity $user.Manager -Server $Domain).Name
                        } catch {
                            $mgrName = "Manager Not Found in $Domain"
                        }
                    }

                    # Map attributes to your specific naming convention
                    [PSCustomObject]@{
                        "First Name"                    = $user.GivenName
                        "Last Name"                     = $user.Surname
                        "Display Name"                  = $user.DisplayName
                        "Employee Name"                 = $user.DisplayName
                        "Title"                         = $user.Title
                        "User Logon Name (UPN)"         = $user.UserPrincipalName
                        "SamAccountName"                = $user.SamAccountName
                        "Email Address"                 = $user.EmailAddress
                        "Department"                    = $user.Department
                        "Employee ID"                   = $user.EmployeeID
                        "Employee Number"               = $user.EmployeeNumber
                        "Employee Type"                 = $user.EmployeeType
                        "Description"                   = $user.Description
                        "Manager Name"                  = $mgrName
                        "Date Created"                  = $user.whenCreated
                        "Password Never Expire"         = $user.PasswordNeverExpires
                        "Reset Password at First Logon" = $user.PasswordExpired
                        "MsExchExtensionAttribute27"    = $user.msExchExtensionAttribute27
                        "MsExchExtensionAttribute28"    = $user.msExchExtensionAttribute28
                        "Custom Attribute 5"            = $user.extensionAttribute5
                        "Custom Attribute 6"            = $user.extensionAttribute6
                        "Distinguished Name"            = $user.DistinguishedName
                        "Domain"                        = $Domain
                    }
                }

                $MasterResults += $CurrentDomainResults
                Write-Host "Successfully added results from $Domain." -ForegroundColor Green
            }
        }
        catch {
            Write-Warning "Could not reach or query domain $Domain. Error: $($_.Exception.Message)"
        }
    }

    if ($MasterResults.Count -eq 0) {
        Write-Warning "No enabled accounts found across the forest."
    } else {
        # 4. Show results in GridView
        $MasterResults | Out-GridView -Title "Forest-Wide Detailed User Report"

        # 5. Export to Desktop
        $Path = "$env:USERPROFILE\Desktop\Detailed_Forest_Users_$(Get-Date -Format 'yyyyMMdd').csv"
        $MasterResults | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
        
        Write-Host "`nDONE: Total of $($MasterResults.Count) accounts exported." -ForegroundColor Green
        Write-Host "Report Path: $Path" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "A critical error occurred: $($_.Exception.Message)"
}