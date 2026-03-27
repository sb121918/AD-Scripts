param(
    [Parameter(Mandatory=$true)]
    [string]$SamAccountName
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $User = Get-ADUser -Identity $SamAccountName -Properties *
    
    if ($null -eq $User) {
        @{ Error = "User not found" } | ConvertTo-Json
        exit
    }
    
    $Groups = @($User.MemberOf | ForEach-Object { 
        $GroupDN = $_
        try {
            $Group = Get-ADGroup -Identity $GroupDN -Properties GroupCategory, GroupScope, mail
            [PSCustomObject]@{
                Name = $Group.Name
                Type = $Group.GroupCategory
                Scope = $Group.GroupScope
                Email = $Group.mail
                DistinguishedName = $Group.DistinguishedName
            }
        } catch {
            [PSCustomObject]@{
                Name = ($GroupDN -split ',')[0].Replace("CN=","")
                Type = "Unknown"
                Scope = "Unknown"
                Email = $null
                DistinguishedName = $GroupDN
            }
        }
    })
    
    $Result = [PSCustomObject]@{
        # Basic Information
        DisplayName       = $User.DisplayName
        SamAccountName    = $User.SamAccountName
        UserPrincipalName = $User.UserPrincipalName
        EmployeeID        = $User.employeeID
        EmployeeNumber    = $User.employeeNumber
        DistinguishedName = $User.DistinguishedName
        
        # Contact Information
        Email             = $User.EmailAddress
        Telephone         = $User.telephoneNumber
        Mobile            = $User.mobile
        ProxyAddresses    = @($User.proxyAddresses)
        
        # Job Information
        Title             = $User.Title
        Department        = $User.Department
        Company           = $User.Company
        Manager           = if ($User.Manager) { ($User.Manager -split ',')[0].Replace("CN=","") } else { $null }
        ManagerDN         = $User.Manager
        
        # Account Status
        Enabled           = $User.Enabled
        LockedOut         = $User.LockedOut
        PasswordExpired   = $User.PasswordExpired
        PasswordLastSet   = $User.PasswordLastSet
        PasswordNeverExpires = $User.PasswordNeverExpires
        LastLogonDate     = $User.LastLogonDate
        
        # Dates
        Created           = $User.WhenCreated
        Modified          = $User.WhenChanged
        AccountExpires    = if ($User.AccountExpirationDate) { $User.AccountExpirationDate } else { "Never" }
        
        # Groups
        GroupCount        = $Groups.Count
        Groups            = $Groups
    }
    
    $Result | ConvertTo-Json -Depth 4
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
