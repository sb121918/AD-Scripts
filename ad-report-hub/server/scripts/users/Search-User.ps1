param(
    [Parameter(Mandatory=$true)]
    [string]$SearchValue,
    [string]$SearchType = "wildcard"
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $WildcardSearch = if ($SearchType -eq "exact") { $SearchValue } else { "*$($SearchValue.Trim('*'))*" }
    
    # Get the Global Catalog Server for forest-wide search
    $CurrentForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $GlobalCatalog = ($CurrentForest.GlobalCatalogs[0]).Name
    
    # Build filter based on search type
    if ($SearchType -eq "exact") {
        $Filter = "mail -eq '$SearchValue' -or UserPrincipalName -eq '$SearchValue'"
    } else {
        $Filter = "Name -like '$WildcardSearch' -or DisplayName -like '$WildcardSearch' -or mail -like '$WildcardSearch' -or SamAccountName -like '$WildcardSearch' -or employeeID -like '$WildcardSearch'"
    }
    
    $ADUsers = Get-ADUser -Filter $Filter -Server "$($GlobalCatalog):3268" -Properties `
        DisplayName, EmailAddress, employeeID, employeeNumber, Title, Department, `
        telephoneNumber, mobile, proxyAddresses, MemberOf, Enabled, Manager, `
        DistinguishedName, UserPrincipalName, WhenCreated, WhenChanged, LastLogonDate
    
    if ($null -eq $ADUsers -or @($ADUsers).Count -eq 0) {
        @() | ConvertTo-Json
        exit
    }
    
    $Results = foreach ($User in $ADUsers) {
        $DomainDN = ($User.DistinguishedName -split 'DC=')[-2..-1] -join '.'
        $Groups = @($User.MemberOf | ForEach-Object { ($_ -split ',')[0].Replace("CN=","") })
        
        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            SamAccountName    = $User.SamAccountName
            UserPrincipalName = $User.UserPrincipalName
            Email             = $User.EmailAddress
            EmployeeID        = $User.employeeID
            EmployeeNumber    = $User.employeeNumber
            Title             = $User.Title
            Department        = $User.Department
            Telephone         = $User.telephoneNumber
            Mobile            = $User.mobile
            Enabled           = $User.Enabled
            Domain            = $DomainDN.Replace(',','')
            DistinguishedName = $User.DistinguishedName
            Manager           = if ($User.Manager) { ($User.Manager -split ',')[0].Replace("CN=","") } else { $null }
            GroupCount        = $Groups.Count
            Groups            = $Groups
            ProxyAddresses    = @($User.proxyAddresses | Where-Object { $_ -like "smtp:*" -or $_ -like "SMTP:*" })
            Created           = $User.WhenCreated
            Modified          = $User.WhenChanged
            LastLogon         = $User.LastLogonDate
        }
    }
    
    @($Results) | ConvertTo-Json -Depth 4
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
