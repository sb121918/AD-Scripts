param(
    [string]$TargetDomain
)

try {
    Import-Module ActiveDirectory -ErrorAction Stop

    if ($TargetDomain) {
        $forest = Get-ADForest -Server $TargetDomain
    } else {
        $forest = Get-ADForest
    }
    $results = @()

    # Forest-wide FSMO roles
    $results += [PSCustomObject]@{
        Role = "Schema Master"
        Holder = $forest.SchemaMaster
        Domain = $forest.RootDomain
        Scope = "Forest"
    }

    $results += [PSCustomObject]@{
        Role = "Domain Naming Master"
        Holder = $forest.DomainNamingMaster
        Domain = $forest.RootDomain
        Scope = "Forest"
    }

    # Domain-specific FSMO roles
    foreach ($domainName in $forest.Domains) {
        try {
            $domain = Get-ADDomain -Identity $domainName -ErrorAction SilentlyContinue

            $results += [PSCustomObject]@{
                Role = "PDC Emulator"
                Holder = $domain.PDCEmulator
                Domain = $domain.DNSRoot
                Scope = "Domain"
            }

            $results += [PSCustomObject]@{
                Role = "RID Master"
                Holder = $domain.RIDMaster
                Domain = $domain.DNSRoot
                Scope = "Domain"
            }

            $results += [PSCustomObject]@{
                Role = "Infrastructure Master"
                Holder = $domain.InfrastructureMaster
                Domain = $domain.DNSRoot
                Scope = "Domain"
            }
        } catch {
            # Skip domains we can't query
        }
    }

    $results | ConvertTo-Json -Depth 5
}
catch {
    @{ Error = $_.Exception.Message } | ConvertTo-Json
    exit 1
}
