try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Get current user identity
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $userName = $currentUser.Name
    $samAccountName = $env:USERNAME
    $computerName = $env:COMPUTERNAME
    
    # Get the user's logon server (domain controller)
    $logonServer = $env:LOGONSERVER -replace '\\\\', ''
    
    # Get domain information
    $userDomain = $env:USERDOMAIN
    $userDnsDomain = $env:USERDNSDOMAIN
    
    # If USERDNSDOMAIN is not set, try to get it from the domain
    if (-not $userDnsDomain) {
        try {
            $domainInfo = Get-ADDomain -Identity $userDomain -ErrorAction SilentlyContinue
            $userDnsDomain = $domainInfo.DNSRoot
        } catch {
            $userDnsDomain = $userDomain
        }
    }
    
    # Try to get AD user details with extended properties
    try {
        # Query AD using the logon server for better reliability
        $serverParam = @{}
        if ($logonServer) {
            $serverParam['Server'] = $logonServer
        }
        
        $adUser = Get-ADUser -Identity $samAccountName @serverParam -Properties `
            DisplayName, GivenName, Surname, EmailAddress, Title, Department, `
            Company, Manager, Office, OfficePhone, MobilePhone, ipPhone, `
            StreetAddress, City, State, St, PostalCode, Country, Co, `
            Description, EmployeeID, EmployeeNumber, `
            WhenCreated, LastLogonDate, PasswordLastSet, PasswordExpired, `
            Enabled, LockedOut, AccountExpirationDate, `
            DistinguishedName, UserPrincipalName, MemberOf, `
            physicalDeliveryOfficeName, l, c
        
        # Get manager name if exists
        $managerName = $null
        if ($adUser.Manager) {
            try {
                $managerObj = Get-ADUser -Identity $adUser.Manager @serverParam -Properties DisplayName
                $managerName = if ($managerObj.DisplayName) { $managerObj.DisplayName } else { $managerObj.Name }
            } catch {
                # Extract CN from DN for manager
                if ($adUser.Manager -match 'CN=([^,]+)') {
                    $managerName = $matches[1]
                } else {
                    $managerName = $adUser.Manager
                }
            }
        }
        
        # Count group memberships
        $groupCount = if ($adUser.MemberOf) { @($adUser.MemberOf).Count } else { 0 }
        
        # Get computer details from AD
        $computerDN = $null
        $computerOS = $null
        try {
            $adComputer = Get-ADComputer -Identity $computerName @serverParam -Properties OperatingSystem, DistinguishedName -ErrorAction SilentlyContinue
            if ($adComputer) {
                $computerDN = $adComputer.DistinguishedName
                $computerOS = $adComputer.OperatingSystem
            }
        } catch {
            # Ignore computer lookup errors
        }
        
        # Resolve location fields (AD uses different attribute names)
        $cityValue = if ($adUser.City) { $adUser.City } elseif ($adUser.l) { $adUser.l } else { $null }
        $stateValue = if ($adUser.State) { $adUser.State } elseif ($adUser.St) { $adUser.St } else { $null }
        $countryValue = if ($adUser.Country) { $adUser.Country } elseif ($adUser.Co) { $adUser.Co } elseif ($adUser.c) { $adUser.c } else { $null }
        $officeValue = if ($adUser.Office) { $adUser.Office } elseif ($adUser.physicalDeliveryOfficeName) { $adUser.physicalDeliveryOfficeName } else { $null }
        $phoneValue = if ($adUser.OfficePhone) { $adUser.OfficePhone } elseif ($adUser.ipPhone) { $adUser.ipPhone } else { $null }
        
        $result = [PSCustomObject]@{
            # Basic Info
            username = $samAccountName
            displayName = if ($adUser.DisplayName) { $adUser.DisplayName } elseif ($adUser.GivenName -or $adUser.Surname) { "$($adUser.GivenName) $($adUser.Surname)".Trim() } else { $samAccountName }
            firstName = $adUser.GivenName
            lastName = $adUser.Surname
            domain = $userDomain
            domainDns = $userDnsDomain
            email = $adUser.EmailAddress
            upn = $adUser.UserPrincipalName
            
            # Job Info
            title = $adUser.Title
            department = $adUser.Department
            company = $adUser.Company
            manager = $managerName
            employeeId = if ($adUser.EmployeeID) { $adUser.EmployeeID } elseif ($adUser.EmployeeNumber) { $adUser.EmployeeNumber } else { $null }
            description = $adUser.Description
            
            # Contact Info
            office = $officeValue
            phone = $phoneValue
            mobile = $adUser.MobilePhone
            
            # Location
            address = $adUser.StreetAddress
            city = $cityValue
            state = $stateValue
            postalCode = $adUser.PostalCode
            country = $countryValue
            
            # Account Info
            computer = $computerName
            computerDN = $computerDN
            computerOS = $computerOS
            logonServer = $logonServer
            fullName = $userName
            distinguishedName = $adUser.DistinguishedName
            enabled = $adUser.Enabled
            lockedOut = $adUser.LockedOut
            passwordExpired = $adUser.PasswordExpired
            
            # Dates
            created = if ($adUser.WhenCreated) { $adUser.WhenCreated.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
            lastLogon = if ($adUser.LastLogonDate) { $adUser.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
            passwordLastSet = if ($adUser.PasswordLastSet) { $adUser.PasswordLastSet.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
            accountExpires = if ($adUser.AccountExpirationDate) { $adUser.AccountExpirationDate.ToString("yyyy-MM-dd HH:mm:ss") } else { $null }
            
            # Groups
            groupCount = $groupCount
            
            # Debug info
            queryServer = $logonServer
        }
    } catch {
        $errorMsg = $_.Exception.Message
        
        # If AD lookup fails, return basic info with error details
        $result = [PSCustomObject]@{
            username = $samAccountName
            displayName = $samAccountName
            firstName = $null
            lastName = $null
            domain = $userDomain
            domainDns = $userDnsDomain
            email = $null
            upn = "$samAccountName@$userDnsDomain"
            title = $null
            department = $null
            company = $null
            manager = $null
            employeeId = $null
            description = $null
            office = $null
            phone = $null
            mobile = $null
            address = $null
            city = $null
            state = $null
            postalCode = $null
            country = $null
            computer = $computerName
            computerDN = $null
            computerOS = $null
            logonServer = $logonServer
            fullName = $userName
            distinguishedName = $null
            enabled = $true
            lockedOut = $false
            passwordExpired = $false
            created = $null
            lastLogon = $null
            passwordLastSet = $null
            accountExpires = $null
            groupCount = 0
            queryServer = $logonServer
            adLookupError = $errorMsg
        }
    }
    
    $result | ConvertTo-Json -Depth 3
}
catch {
    # Ultimate fallback
    [PSCustomObject]@{
        username = $env:USERNAME
        displayName = $env:USERNAME
        firstName = $null
        lastName = $null
        domain = $env:USERDOMAIN
        domainDns = $env:USERDNSDOMAIN
        email = $null
        computer = $env:COMPUTERNAME
        logonServer = ($env:LOGONSERVER -replace '\\\\', '')
        fullName = "$env:USERDOMAIN\$env:USERNAME"
        fallbackError = $_.Exception.Message
    } | ConvertTo-Json
}
