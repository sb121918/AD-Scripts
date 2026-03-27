param([string]$TargetDomain)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    
    $serverParam = @{}
    if ($TargetDomain) { $serverParam['Server'] = $TargetDomain }
    
    $AllDomains = if ($TargetDomain) { @($TargetDomain) } else { (Get-ADForest @serverParam).Domains }
    $FullReport = New-Object System.Collections.Generic.List[PSObject]

    # Fetch all password policies (no limit)
    foreach ($Domain in $AllDomains) {
        # Get Default Domain Password Policy
        try {
            $DefaultPolicy = Get-ADDefaultDomainPasswordPolicy -Server $Domain
            if ($DefaultPolicy) {
                $FullReport.Add([PSCustomObject]@{
                    Domain                    = $Domain
                    PolicyName                = "DEFAULT DOMAIN POLICY"
                    Type                      = "Default"
                    Precedence                = "N/A"
                    Complexity                = $DefaultPolicy.ComplexityEnabled
                    MinPasswordLength         = $DefaultPolicy.MinPasswordLength
                    PasswordHistory           = $DefaultPolicy.PasswordHistoryCount
                    LockoutThreshold          = $DefaultPolicy.LockoutThreshold
                    LockoutDuration           = if ($DefaultPolicy.LockoutDuration) { $DefaultPolicy.LockoutDuration.ToString() } else { "N/A" }
                    LockoutObservationWindow  = if ($DefaultPolicy.LockoutObservationWindow) { $DefaultPolicy.LockoutObservationWindow.ToString() } else { "N/A" }
                    MaxPasswordAgeDays        = $DefaultPolicy.MaxPasswordAge.Days
                    MinPasswordAgeDays        = $DefaultPolicy.MinPasswordAge.Days
                    ReversibleEncryption      = $DefaultPolicy.ReversibleEncryptionEnabled
                    AppliesTo                 = "All Users (Domain)"
                    AppliesToCount            = "All"
                })
            }
        } catch {
            # Skip if no default policy
        }

        # Get Fine-Grained Password Policies (PSOs)
        try {
            $FGPPs = Get-ADFineGrainedPasswordPolicy -Filter * -Server $Domain -Properties *
            
            if ($FGPPs) {
                foreach ($Policy in $FGPPs) {
                    # Resolve AppliesTo DNs to readable names
                    $AppliesToNames = @()
                    $AppliesToCount = 0
                    
                    if ($Policy.AppliesTo) {
                        $AppliesToCount = $Policy.AppliesTo.Count
                        foreach ($dn in $Policy.AppliesTo) {
                            try {
                                # Try to get the object name from DN
                                $obj = Get-ADObject -Identity $dn -Server $Domain -Properties Name, ObjectClass
                                $objType = switch ($obj.ObjectClass) {
                                    "user" { "User" }
                                    "group" { "Group" }
                                    "inetOrgPerson" { "User" }
                                    default { $obj.ObjectClass }
                                }
                                $AppliesToNames += "$($obj.Name) ($objType)"
                            } catch {
                                # If can't resolve, extract CN from DN
                                if ($dn -match "^CN=([^,]+),") {
                                    $AppliesToNames += $Matches[1]
                                } else {
                                    $AppliesToNames += $dn
                                }
                            }
                        }
                    }
                    
                    $FullReport.Add([PSCustomObject]@{
                        Domain                    = $Domain
                        PolicyName                = $Policy.Name
                        Type                      = "Fine-Grained (PSO)"
                        Precedence                = $Policy.Precedence
                        Complexity                = $Policy.ComplexityEnabled
                        MinPasswordLength         = $Policy.MinPasswordLength
                        PasswordHistory           = $Policy.PasswordHistoryCount
                        LockoutThreshold          = $Policy.LockoutThreshold
                        LockoutDuration           = if ($Policy.LockoutDuration) { $Policy.LockoutDuration.ToString() } else { "N/A" }
                        LockoutObservationWindow  = if ($Policy.LockoutObservationWindow) { $Policy.LockoutObservationWindow.ToString() } else { "N/A" }
                        MaxPasswordAgeDays        = $Policy.MaxPasswordAge.Days
                        MinPasswordAgeDays        = $Policy.MinPasswordAge.Days
                        ReversibleEncryption      = $Policy.ReversibleEncryptionEnabled
                        AppliesTo                 = if ($AppliesToNames.Count -gt 0) { ($AppliesToNames -join '; ') } else { "Not Applied" }
                        AppliesToCount            = $AppliesToCount
                    })
                }
            }
        } catch {
            # Skip if error getting FGPPs
        }
    }

    @($FullReport) | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
