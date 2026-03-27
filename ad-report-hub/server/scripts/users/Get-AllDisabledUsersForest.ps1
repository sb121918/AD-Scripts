param()

$ErrorActionPreference = 'SilentlyContinue'

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$AllUsers = @()

try {
    $Forest = Get-ADForest -ErrorAction Stop

    foreach ($DomainName in $Forest.Domains) {
        try {
            $Users = Get-ADUser -Filter { Enabled -eq $false } -Server $DomainName -Properties DisplayName, EmailAddress, Department, Title, WhenCreated, WhenChanged -ErrorAction SilentlyContinue
            
            foreach ($User in $Users) {
                $AllUsers += [PSCustomObject]@{
                    DisplayName       = $User.DisplayName
                    SamAccountName    = $User.SamAccountName
                    Email             = $User.EmailAddress
                    Department        = $User.Department
                    Title             = $User.Title
                    Enabled           = $false
                    Domain            = $DomainName
                    Created           = $User.WhenCreated
                    Modified          = $User.WhenChanged
                    DistinguishedName = $User.DistinguishedName
                }
            }
        }
        catch { }
    }
}
catch {
    @{ Error = "Failed to query forest: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($AllUsers) | ConvertTo-Json -Depth 3
