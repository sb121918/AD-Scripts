param(
    [Parameter(Mandatory=$true)]
    [string]$LDAPFilter,
    [string]$ObjectTypes = "user,computer,group",
    [int]$PageSize = 25,
    [int]$Page = 1
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $Types = $ObjectTypes.Split(',').Trim()
    $Skip = ($Page - 1) * $PageSize
    
    $Results = @{
        Users = @()
        Computers = @()
        Groups = @()
        Contacts = @()
        Printers = @()
        OUs = @()
        GPOs = @()
        Pagination = @{
            Page = $Page
            PageSize = $PageSize
            TotalUsers = 0
            TotalComputers = 0
            TotalGroups = 0
            TotalContacts = 0
            TotalPrinters = 0
            TotalOUs = 0
            TotalGPOs = 0
        }
    }
    
    # Search Users - filter already includes objectClass from query builder
    if ($Types -contains "user") {
        $AllUsers = @(Get-ADUser -LDAPFilter $LDAPFilter -Properties DisplayName, mail, Department, Enabled -ErrorAction SilentlyContinue)
        $Results.Pagination.TotalUsers = $AllUsers.Count
        
        $PagedUsers = $AllUsers | Select-Object -Skip $Skip -First $PageSize
        $Results.Users = @($PagedUsers | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.DisplayName
                SamAccountName = $_.SamAccountName
                Email = $_.mail
                Department = $_.Department
                Enabled = $_.Enabled
                ObjectType = "User"
            }
        })
    }
    
    # Search Computers
    if ($Types -contains "computer") {
        $AllComputers = @(Get-ADComputer -LDAPFilter $LDAPFilter -Properties OperatingSystem, Enabled -ErrorAction SilentlyContinue)
        $Results.Pagination.TotalComputers = $AllComputers.Count
        
        $PagedComputers = $AllComputers | Select-Object -Skip $Skip -First $PageSize
        $Results.Computers = @($PagedComputers | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                OperatingSystem = $_.OperatingSystem
                Enabled = $_.Enabled
                ObjectType = "Computer"
            }
        })
    }
    
    # Search Groups
    if ($Types -contains "group") {
        $AllGroups = @(Get-ADGroup -LDAPFilter $LDAPFilter -Properties GroupCategory, mail -ErrorAction SilentlyContinue)
        $Results.Pagination.TotalGroups = $AllGroups.Count
        
        $PagedGroups = $AllGroups | Select-Object -Skip $Skip -First $PageSize
        $Results.Groups = @($PagedGroups | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Type = [string]$_.GroupCategory
                Email = $_.mail
                ObjectType = "Group"
            }
        })
    }
    
    # Search Contacts
    if ($Types -contains "contact") {
        $AllContacts = @(Get-ADObject -LDAPFilter $LDAPFilter -Properties DisplayName, mail, company, department -ErrorAction SilentlyContinue | Where-Object { $_.ObjectClass -eq 'contact' })
        $Results.Pagination.TotalContacts = $AllContacts.Count
        
        $PagedContacts = $AllContacts | Select-Object -Skip $Skip -First $PageSize
        $Results.Contacts = @($PagedContacts | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.DisplayName
                Email = $_.mail
                Company = $_.company
                Department = $_.department
                ObjectType = "Contact"
            }
        })
    }
    
    # Search Printers
    if ($Types -contains "printer") {
        $AllPrinters = @(Get-ADObject -LDAPFilter $LDAPFilter -Properties serverName, location, description -ErrorAction SilentlyContinue | Where-Object { $_.ObjectClass -eq 'printQueue' })
        $Results.Pagination.TotalPrinters = $AllPrinters.Count
        
        $PagedPrinters = $AllPrinters | Select-Object -Skip $Skip -First $PageSize
        $Results.Printers = @($PagedPrinters | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Server = $_.serverName
                Location = $_.location
                ObjectType = "Printer"
            }
        })
    }
    
    # Search OUs
    if ($Types -contains "ou") {
        $AllOUs = @(Get-ADObject -LDAPFilter $LDAPFilter -Properties description -ErrorAction SilentlyContinue | Where-Object { $_.ObjectClass -eq 'organizationalUnit' })
        $Results.Pagination.TotalOUs = $AllOUs.Count
        
        $PagedOUs = $AllOUs | Select-Object -Skip $Skip -First $PageSize
        $Results.OUs = @($PagedOUs | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Description = $_.description
                DistinguishedName = $_.DistinguishedName
                ObjectType = "OrganizationalUnit"
            }
        })
    }
    
    # Search GPOs
    if ($Types -contains "gpo") {
        $AllGPOs = @(Get-ADObject -LDAPFilter $LDAPFilter -Properties displayName, whenCreated -ErrorAction SilentlyContinue | Where-Object { $_.ObjectClass -eq 'groupPolicyContainer' })
        $Results.Pagination.TotalGPOs = $AllGPOs.Count
        
        $PagedGPOs = $AllGPOs | Select-Object -Skip $Skip -First $PageSize
        $Results.GPOs = @($PagedGPOs | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.displayName
                Status = "Enabled"
                Created = if ($_.whenCreated) { $_.whenCreated.ToString('yyyy-MM-dd') } else { '-' }
                ObjectType = "GPO"
            }
        })
    }

    $Results | ConvertTo-Json -Depth 4
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
