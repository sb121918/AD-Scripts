param(
    [Parameter(Mandatory=$true)]
    [string]$SearchValue,
    [string]$ObjectTypes = "user,computer,group",
    [int]$PageSize = 25,
    [int]$Page = 1
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $WildcardSearch = "*$($SearchValue.Trim('*'))*"
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
    
    # Search Users
    if ($Types -contains "user") {
        $Filter = "Name -like '$WildcardSearch' -or DisplayName -like '$WildcardSearch' -or mail -like '$WildcardSearch' -or SamAccountName -like '$WildcardSearch'"
        $AllUsers = @(Get-ADUser -Filter $Filter -Properties DisplayName, mail, Department, Enabled)
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
        $Filter = "Name -like '$WildcardSearch'"
        $AllComputers = @(Get-ADComputer -Filter $Filter -Properties OperatingSystem, Enabled)
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
        $Filter = "Name -like '$WildcardSearch' -or DisplayName -like '$WildcardSearch'"
        $AllGroups = @(Get-ADGroup -Filter $Filter -Properties GroupCategory, mail)
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
        $Filter = "(objectClass=contact)(|(name=$WildcardSearch)(displayName=$WildcardSearch)(mail=$WildcardSearch))"
        $LdapFilter = "(&$Filter)"
        $AllContacts = @(Get-ADObject -LDAPFilter $LdapFilter -Properties DisplayName, mail, company, department, telephoneNumber -ErrorAction SilentlyContinue)
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
        $LdapFilter = "(&(objectClass=printQueue)(|(name=$WildcardSearch)(serverName=$WildcardSearch)(location=$WildcardSearch)))"
        $AllPrinters = @(Get-ADObject -LDAPFilter $LdapFilter -Properties serverName, location, description -ErrorAction SilentlyContinue)
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
        $LdapFilter = "(&(objectClass=organizationalUnit)(|(name=$WildcardSearch)(description=$WildcardSearch)))"
        $AllOUs = @(Get-ADObject -LDAPFilter $LdapFilter -Properties description -ErrorAction SilentlyContinue)
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
        $LdapFilter = "(&(objectClass=groupPolicyContainer)(|(displayName=$WildcardSearch)(name=$WildcardSearch)))"
        $AllGPOs = @(Get-ADObject -LDAPFilter $LdapFilter -Properties displayName, gPCFileSysPath, whenCreated -ErrorAction SilentlyContinue)
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
