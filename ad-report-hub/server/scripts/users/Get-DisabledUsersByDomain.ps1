param(
    [Parameter(Mandatory=$true)]
    [string]$TargetDomain
)

$ErrorActionPreference = 'SilentlyContinue'

try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

$Results = @()

try {
    # Fetch all disabled users from specific domain
    $Users = Get-ADUser -Filter { Enabled -eq $false } -Server $TargetDomain -Properties DisplayName, EmailAddress, Department, Title, WhenCreated, WhenChanged -ErrorAction Stop |
             Sort-Object DisplayName
    
    foreach ($User in $Users) {
        $Results += [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            Email             = $User.EmailAddress
            Department        = $User.Department
            Title             = $User.Title
            Enabled           = $false
            Domain            = $TargetDomain
            Created           = $User.WhenCreated
            Modified          = $User.WhenChanged
            DistinguishedName = $User.DistinguishedName
        }
    }
}
catch {
    @{ Error = "Failed to query disabled users for ${TargetDomain}: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

@($Results) | ConvertTo-Json -Depth 3
