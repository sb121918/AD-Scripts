<#
.SYNOPSIS
    Identifies an Active Directory object by its ObjectGUID or SID.

.DESCRIPTION
    Looks up a single AD object (user, group, computer, OU, etc.) using either
    its ObjectGUID or Security Identifier (SID). Returns core properties;
    use -IncludeAllProperties for full attribute list.

.PARAMETER Identity
    The ObjectGUID or SID of the AD object.
    SID format:  "S-1-5-21-4186143834-2626045635-1021053583-11850"
    GUID format: "a1b2c3d4-e5f6-7890-abcd-ef1234567890" or "a1b2c3d4e5f67890abcdef1234567890"

.PARAMETER Server
    Optional. Domain name or DC to query (e.g. "contoso.com" or "DC01.contoso.com").
    If omitted, uses the default domain for the current user.

.PARAMETER IncludeAllProperties
    If set, returns all LDAP attributes; otherwise returns a concise set of properties.

.EXAMPLE
    .\Get-ADObjectByGUID.ps1 -Identity "S-1-5-21-4186143834-2626045635-1021053583-11850"
.EXAMPLE
    .\Get-ADObjectByGUID.ps1 -Identity "a1b2c3d4-e5f6-7890-abcd-ef1234567890" -Server "contoso.com" -IncludeAllProperties
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$Identity,

    [string]$Server = "",

    [switch]$IncludeAllProperties
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $ErrorActionPreference = 'Stop'

    $inputVal = $Identity.Trim()

    # Detect SID (e.g. S-1-5-21-4186143834-2626045635-1021053583-11850) vs GUID
    $isSid = $inputVal -match '^S-1-\d+(-\d+)+$'
    if ($isSid) {
        $lookupIdentity = $inputVal
    } else {
        # Normalize GUID: convert 32-char hex to 8-4-4-4-12 format
        $guidStr = $inputVal -replace '\s+', ''
        if ($guidStr -match '^([0-9a-fA-F]{8})([0-9a-fA-F]{4})([0-9a-fA-F]{4})([0-9a-fA-F]{4})([0-9a-fA-F]{12})$') {
            $guidStr = "$($matches[1])-$($matches[2])-$($matches[3])-$($matches[4])-$($matches[5])"
        }
        if ($guidStr -notmatch '^\{?[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}?$') {
            @{ Error = "Invalid Identity. Use SID (e.g. S-1-5-21-...) or GUID (e.g. a1b2c3d4-e5f6-7890-abcd-ef1234567890). Got: $Identity" } | ConvertTo-Json
            exit 1
        }
        $lookupIdentity = $guidStr
    }

    $params = @{
        Identity    = $lookupIdentity
        ErrorAction = 'Stop'
    }
    if ($Server) { $params['Server'] = $Server }
    if ($IncludeAllProperties) {
        $params['Properties'] = '*'
    } else {
        $params['Properties'] = 'WhenCreated', 'WhenChanged', 'ObjectSid'
    }

    $obj = Get-ADObject @params

    if (-not $obj) {
        @{ Found = $false; Error = "No object found: $Identity" } | ConvertTo-Json
        exit 0
    }

    $objectSidStr = $null
    if ($obj.ObjectSid) { $objectSidStr = $obj.ObjectSid.ToString() }

    if ($IncludeAllProperties) {
        $result = [ordered]@{
            Found             = $true
            DistinguishedName = $obj.DistinguishedName
            Name              = $obj.Name
            ObjectClass       = $obj.ObjectClass
            ObjectGUID       = $obj.ObjectGUID.Guid
            ObjectSid        = $objectSidStr
            AllProperties    = @{}
        }
        foreach ($p in $obj.PSObject.Properties) {
            $val = $p.Value
            if ($val -is [byte[]]) { $val = [System.Convert]::ToBase64String($val) }
            elseif ($val -is [DateTime]) { $val = $val.ToString('o') }
            elseif ($val -is [System.Security.Principal.SecurityIdentifier]) { $val = $val.ToString() }
            $result.AllProperties[$p.Name] = $val
        }
        $result | ConvertTo-Json -Depth 10
    } else {
        $result = [PSCustomObject]@{
            Found             = $true
            DistinguishedName = $obj.DistinguishedName
            Name              = $obj.Name
            ObjectClass       = $obj.ObjectClass
            ObjectGUID       = $obj.ObjectGUID.Guid
            ObjectSid        = $objectSidStr
            WhenCreated      = if ($obj.WhenCreated) { $obj.WhenCreated.ToString('o') } else { $null }
            WhenChanged      = if ($obj.WhenChanged) { $obj.WhenChanged.ToString('o') } else { $null }
        }
        $result | ConvertTo-Json
    }
} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
    @{ Found = $false; Error = "No AD object found: $Identity" } | ConvertTo-Json
    exit 0
} catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
