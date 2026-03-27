param(
    [Parameter(Mandatory=$false)]
    [string]$SearchValue = "",

    [Parameter(Mandatory=$false)]
    [string]$TargetDomain = ""
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $ServerArgs = @{}
    if ($TargetDomain -and $TargetDomain.Trim() -ne "") {
        $ServerArgs.Server = $TargetDomain.Trim()
    }

    function Escape-ADFilterValue([string]$Value) {
        if ($null -eq $Value) { return "" }
        return $Value.Replace("'", "''")
    }

    $Users = @()
    $sv = ($SearchValue || "").Trim()

    if ($sv) {
        # Targeted lookup for a specific user (samAccountName / UPN / mail)
        $candidates = @()
        try {
            $u = Get-ADUser @ServerArgs -Identity $sv -Properties DisplayName, EmailAddress, Department, Title, LockedOut, AccountLockoutTime -ErrorAction Stop
            if ($u) { $candidates = @($u) }
        } catch {
            $escaped = Escape-ADFilterValue $sv
            $Filter = "SamAccountName -eq '$escaped' -or UserPrincipalName -eq '$escaped' -or mail -eq '$escaped'"
            $candidates = @(Get-ADUser @ServerArgs -Filter $Filter -Properties DisplayName, EmailAddress, Department, Title, LockedOut, AccountLockoutTime -ErrorAction SilentlyContinue)
        }

        if ($candidates -and @($candidates).Count -gt 0) {
            $Users = @($candidates | Where-Object { $_.LockedOut -eq $true })
        } else {
            $Users = @()
        }
    } else {
        # Fetch all locked out users (no limit)
        $Users = Search-ADAccount @ServerArgs -LockedOut -UsersOnly |
                 Get-ADUser @ServerArgs -Properties DisplayName, EmailAddress, Department, Title, LockedOut, AccountLockoutTime
    }
    
    $Results = foreach ($User in $Users) {
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            Email             = $User.EmailAddress
            Department        = $User.Department
            Title             = $User.Title
            LockedOut         = $User.LockedOut
            LockoutTime       = $User.AccountLockoutTime
            DistinguishedName = $User.DistinguishedName
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 3
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
