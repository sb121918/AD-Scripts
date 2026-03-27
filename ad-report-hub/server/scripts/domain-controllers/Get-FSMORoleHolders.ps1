param(
    [string]$ForestDomain = "",
    [string]$TargetDomain = ""
)

$ErrorActionPreference = 'SilentlyContinue'

try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    @{ Error = "Failed to load ActiveDirectory module: $($_.Exception.Message)" } | ConvertTo-Json
    exit 1
}

try {
    $domainToQuery = if ($TargetDomain) { $TargetDomain } elseif ($ForestDomain) { $ForestDomain } else { $null }

    # Domain FSMO roles
    $domainObj = if ($domainToQuery) { Get-ADDomain -Server $domainToQuery -ErrorAction Stop } else { Get-ADDomain -ErrorAction Stop }
    $domainName = $domainObj.DNSRoot

    $results = @(
        [PSCustomObject]@{ Role = "PDC Emulator";          Holder = $domainObj.PDCEmulator;          Domain = $domainName; Scope = "Domain" }
        [PSCustomObject]@{ Role = "RID Master";            Holder = $domainObj.RIDMaster;            Domain = $domainName; Scope = "Domain" }
        [PSCustomObject]@{ Role = "Infrastructure Master"; Holder = $domainObj.InfrastructureMaster; Domain = $domainName; Scope = "Domain" }
    )

    # Forest FSMO roles (best-effort; may fail in trusted forest scenarios)
    try {
        $forestAnchor = if ($ForestDomain) { $ForestDomain } elseif ($domainToQuery) { $domainToQuery } else { $null }
        $forestObj = if ($forestAnchor) { Get-ADForest -Server $forestAnchor -ErrorAction Stop } else { Get-ADForest -ErrorAction Stop }

        $results += [PSCustomObject]@{ Role = "Schema Master";        Holder = $forestObj.SchemaMaster;        Domain = $forestObj.RootDomain; Scope = "Forest" }
        $results += [PSCustomObject]@{ Role = "Domain Naming Master"; Holder = $forestObj.DomainNamingMaster;  Domain = $forestObj.RootDomain; Scope = "Forest" }
    } catch {
        # Ignore forest role lookup errors
    }

    @($results) | ConvertTo-Json -Depth 4 -Compress
} catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}

