#Requires -Version 7.2

<#
.SYNOPSIS
Creates an Azure regional migration inventory and Excel workbook using Az PowerShell.
.DESCRIPTION
Self-contained, read-only collector. Requires Az.Accounts, Az.ResourceGraph and
Az.Resources. Queries the current/selected tenant in pages of up to 1000 resources,
filters for regional impact, normalizes service-specific SKUs and exports
an Excel workbook plus JSONL companions. No companion script, Azure CLI or Excel
installation is required.
.PARAMETER SourceRegion
Canonical Azure source region, such as canadacentral or southeastasia.
.PARAMETER TargetRegion
Optional destination for planning context only; does not validate deployment readiness.
.PARAMETER SubscriptionId
Optional subscription IDs queried together. If omitted, uses tenant-wide scope.
All supplied subscriptions must be visible in the selected/current tenant.
Non-Enabled subscriptions remain in query scope and are reported as warnings.
.PARAMETER TenantId
Optional tenant scope. Defaults to the current Az context's tenant, not every signed-in tenant.
.PARAMETER OutputDirectory
New output folder. Defaults to RegionalInventory-<timestamp> under the current directory.
.PARAMETER Login
Signs in with Connect-AzAccount. Omit when an Az PowerShell session already exists.
.PARAMETER UseDeviceCode
Uses device authentication with Login, subject to organizational policy.
.PARAMETER CheckOnly
Validates prerequisites and live read access without collecting inventory files.
.PARAMETER RowsPerSheet
Maximum data rows per worksheet. Default 250000; summary must fit on one sheet.
.EXAMPLE
.\Export-AzureRegionalInventory.ps1 -SourceRegion canadacentral -Login
.EXAMPLE
.\Export-AzureRegionalInventory.ps1 -SourceRegion southeastasia -TenantId '<tenant-guid>' -SubscriptionId '<subscription-guid>' -CheckOnly
.EXAMPLE
Get-Help .\Export-AzureRegionalInventory.ps1 -Full
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^[a-z0-9]+$')][string]$SourceRegion,
    [ValidatePattern('^[a-z0-9]*$')][string]$TargetRegion = '',
    [ValidateNotNullOrEmpty()][string[]]$SubscriptionId,
    [string]$TenantId,
    [string]$OutputDirectory = (Join-Path $PWD ("RegionalInventory-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))),
    [switch]$Login,
    [switch]$UseDeviceCode,
    [switch]$CheckOnly,
    [ValidateRange(1, 1048575)][int]$RowsPerSheet = 250000
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region Azure prerequisites and paginated queries
function Import-InventoryAzModules {
    foreach ($name in @('Az.Accounts', 'Az.ResourceGraph', 'Az.Resources')) {
        if (Get-Module -Name $name -Verbose:$false) { continue }
        if (-not (Get-Module -ListAvailable -Name $name -Verbose:$false)) {
            throw "Required module $name is missing. Install it using your approved method, for example: Install-Module $name -Scope CurrentUser"
        }
        Import-Module $name -ErrorAction Stop -Verbose:$false
    }
    $required = @{
        'Search-AzGraph' = @('Query', 'Subscription', 'UseTenantScope', 'First', 'SkipToken', 'DefaultProfile')
        'Connect-AzAccount' = @('Tenant', 'UseDeviceAuthentication', 'Scope')
        'Get-AzContext' = @()
        'Set-AzContext' = @('Subscription', 'Tenant', 'Context', 'Scope')
        'Get-AzSubscription' = @('SubscriptionId', 'TenantId', 'DefaultProfile')
        'Get-AzLocation' = @('DefaultProfile')
    }
    foreach ($name in $required.Keys) {
        $command = Get-Command $name -ErrorAction Stop
        foreach ($parameter in $required[$name]) {
            if (-not $command.Parameters.ContainsKey($parameter)) {
                throw "$name is missing -$parameter. Update the owning Az module using your approved method."
            }
        }
    }
    foreach ($name in @('Az.Accounts', 'Az.ResourceGraph', 'Az.Resources')) {
        $module = Get-Module $name
        [pscustomobject]@{ Name = $module.Name; Version = $module.Version.ToString() }
    }
}

function Invoke-InventoryGraph {
    param(
        [Parameter(Mandatory)][string]$Query,
        [string[]]$SubscriptionId,
        [ValidateRange(1, 1000)][int]$First = 1000,
        [string]$SkipToken,
        $DefaultProfile,
        [ValidateRange(0, 12)][int]$MaxRetries = 6
    )
    $parameters = @{ Query = $Query; First = $First; ErrorAction = 'Stop'; Verbose = $false }
    if ($SubscriptionId) { $parameters.Subscription = @($SubscriptionId) }
    else { $parameters.UseTenantScope = $true }
    $scopeLabel = if ($SubscriptionId) { "Subscriptions ($($SubscriptionId.Count))" } else { 'Tenant' }
    if ($SkipToken) { $parameters.SkipToken = $SkipToken }
    if ($null -ne $DefaultProfile) { $parameters.DefaultProfile = $DefaultProfile }
    for ($attempt = 0; ; $attempt++) {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        try {
            # Keep the response object: enumerating it first would lose SkipToken metadata.
            $page = Search-AzGraph @parameters
            if ($null -eq $page -or $null -eq $page.PSObject.Properties['Data'] -or
                $null -eq $page.PSObject.Properties['SkipToken']) {
                throw 'Search-AzGraph returned an invalid response; Data and SkipToken are required.'
            }
            $rows = ConvertTo-Json -InputObject @($page.Data) -Depth 100 -Compress |
                ConvertFrom-Json -AsHashtable -Depth 100 -NoEnumerate
            return @{ data = $rows; skipToken = $page.SkipToken }
        }
        catch {
            $errorText = $_.Exception.Message
            $retryable = $errorText -match '(?i)429|TooManyRequests|throttl|502|503|504|InternalServerError|ServiceUnavailable|GatewayTimeout|timed out|connection reset'
            if (-not $retryable -or $attempt -ge $MaxRetries) { throw }
            $delay = [Math]::Min(120, [Math]::Pow(2, $attempt + 1) + (Get-Random -Minimum 0 -Maximum 4))
            Write-Warning "Transient ARG failure; retry $($attempt + 1)/$MaxRetries in $delay seconds: $errorText"
            Start-Sleep -Seconds $delay
        }
        finally {
            $timer.Stop()
            Write-Verbose ("Search-AzGraph [{0}]: {1:N2}s (attempt {2})" -f $scopeLabel, $timer.Elapsed.TotalSeconds, ($attempt + 1))
        }
    }
}

function Get-ArgCount {
    param([string[]]$SubscriptionId, [string]$BaseQuery, $DefaultProfile)
    $response = Invoke-InventoryGraph -SubscriptionId $SubscriptionId -DefaultProfile $DefaultProfile `
        -First 1 -Query "$BaseQuery | summarize ResourceCount=count()"
    if (@($response.data).Count -ne 1 -or -not $response.data[0].ContainsKey('ResourceCount')) {
        throw 'ARG count response is invalid for the requested query scope.'
    }
    return [long]$response.data[0].ResourceCount
}

function Get-RegionalResourceQuery {
    param([Parameter(Mandatory)][ValidatePattern('^[a-z0-9]+$')][string]$SourceRegion)
    return "Resources | where location =~ '$SourceRegion' or location =~ 'global' or isempty(location) or replace_string(tolower(tostring(properties.locations)), ' ', '') contains '$SourceRegion' or properties.primaryLocation =~ '$SourceRegion' or properties.secondaryLocation =~ '$SourceRegion'"
}

function Invoke-ArgPaged {
    param(
        [string[]]$SubscriptionId,
        [Parameter(Mandatory)][string]$BaseQuery,
        [Parameter(Mandatory)][string]$Projection,
        [Parameter(Mandatory)][scriptblock]$OnPage,
        [ValidateRange(1, 1000)][int]$PageSize = 1000,
        [ValidateRange(0, [long]::MaxValue)][long]$InitialCount,
        $DefaultProfile,
        [string]$DatasetName = 'Resources'
    )
    $expected = if ($PSBoundParameters.ContainsKey('InitialCount')) { $InitialCount }
        else { Get-ArgCount -SubscriptionId $SubscriptionId -BaseQuery $BaseQuery -DefaultProfile $DefaultProfile }
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $tokens = [Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $token = $null
    $pages = 0
    $scopeLabel = if ($SubscriptionId) { "Subscriptions ($($SubscriptionId.Count))" } else { 'Tenant' }
    $query = "$BaseQuery | project $Projection | order by id asc"
    do {
        $response = Invoke-InventoryGraph -SubscriptionId $SubscriptionId -DefaultProfile $DefaultProfile `
            -First $PageSize -Query $query -SkipToken $token
        if (-not $response.ContainsKey('data')) { throw "ARG response is missing data for $scopeLabel." }
        $rows = @($response.data)
        $pages++
        foreach ($row in $rows) {
            if (-not $row.id) { throw 'Every paged query must project id.' }
            if (-not $seen.Add([string]$row.id)) {
                throw "Duplicate resource ID while paging: $($row.id). Inventory may have changed; rerun."
            }
        }
        if ($rows.Count -gt 0) { & $OnPage $rows }
        $token = $response['skipToken']
        if ($token -and -not $tokens.Add([string]$token)) { throw 'ARG returned a repeated continuation token.' }
        if ($token -and $rows.Count -eq 0) { throw 'ARG returned an empty page with a continuation token.' }
        Write-Host "  Retrieved $DatasetName page $pages ($($rows.Count) row(s)); total so far: $($seen.Count)/$expected [$scopeLabel]" -ForegroundColor DarkGray
    } while ($token)
    $after = Get-ArgCount -SubscriptionId $SubscriptionId -BaseQuery $BaseQuery -DefaultProfile $DefaultProfile
    if ($seen.Count -ne $expected -or $after -ne $expected) {
        throw "ARG count mismatch for $scopeLabel : before=$expected, downloaded=$($seen.Count), after=$after. Rerun in a quieter window."
    }
    [pscustomobject]@{ Expected = $expected; Downloaded = $seen.Count; Pages = $pages; After = $after }
}

#endregion

#region Service metadata and inventory helpers
function Get-Field {
    param($Object, [string]$Path)
    $current = $Object
    foreach ($part in $Path.Split('.')) {
        if ($current -isnot [Collections.IDictionary] -or -not $current.Contains($part)) { return $null }
        $current = $current[$part]
    }
    return ,$current
}

function ConvertTo-CompactJson {
    param($Value)
    if ($null -eq $Value) { return '' }
    return ConvertTo-Json -InputObject $Value -Depth 100 -Compress
}

function Get-ServiceProfile {
    param([Collections.IDictionary]$Resource, [Collections.Generic.Dictionary[string, object]]$Index)
    $p = $Resource['properties']
    $type = ([string]$Resource.type).ToLowerInvariant()
    $sku = Get-Field $Resource 'sku.name'
    $tier = Get-Field $Resource 'sku.tier'
    $capacity = Get-Field $Resource 'sku.capacity'
    $basis = 'Resource SKU'
    $source = [string]$Resource.id
    $details = [ordered]@{}
    $guidance = 'Validate service-specific regional support, data migration, identity, networking, quotas, and cutover.'
    switch -Regex ($type) {
        '^microsoft.compute/(virtualmachines|virtualmachinescalesets)$' {
            if ($type -eq 'microsoft.compute/virtualmachines') {
                $sku = Get-Field $p 'hardwareProfile.vmSize'
                $capacity = 1
            }
            $basis = 'VM size; capacity is VM/VMSS instance count'
            foreach ($path in @('storageProfile.osDisk.osType', 'storageProfile.osDisk.diskSizeGB',
                'storageProfile.osDisk.managedDisk.storageAccountType', 'storageProfile.dataDisks',
                'storageProfile.imageReference', 'priority', 'licenseType',
                'availabilitySet.id', 'host.id', 'securityProfile.securityType')) {
                # Keep disk/image metadata, never VM extensions or customData.
                $readPath = if ($type -eq 'microsoft.compute/virtualmachinescalesets' -and $path -notin @('availabilitySet.id', 'host.id')) {
                    "virtualMachineProfile.$path"
                } else { $path }
                $value = Get-Field $p $readPath
                if ($null -ne $value) { $details[$path] = $value }
            }
            $guidance = 'Validate target VM family quota/capacity, zone/size support, image availability, disks, backup, licensing, and IP changes.'
        }
        '^microsoft.web/sites(/slots)?$' {
            $planId = [string](Get-Field $p 'serverFarmId')
            $basis = 'Inherited from App Service plan'
            $source = $planId
            if ($planId -and $Index.ContainsKey($planId)) {
                $plan = $Index[$planId]
                $sku = $plan.SkuName; $tier = $plan.SkuTier; $capacity = $plan.SkuCapacity
                $details.PlanLocation = $plan.Location
            }
            else { $details.PlanResolution = 'Plan not visible in ARG scan; SKU unknown' }
            foreach ($path in @('state', 'reserved', 'httpsOnly', 'virtualNetworkSubnetId',
                'siteConfig.linuxFxVersion', 'siteConfig.windowsFxVersion', 'siteConfig.minimumElasticInstanceCount')) {
                $details[$path] = Get-Field $p $path
            }
            $guidance = 'Recreate app and target-region plan; migrate content, slots, certificates, identities, VNet integration and DNS. Worker capacity is shared, not per-app.'
        }
        '^microsoft.web/serverfarms$' {
            $basis = 'App Service plan tier/size; capacity is shared workers'
            foreach ($path in @('reserved', 'zoneRedundant', 'maximumElasticWorkerCount', 'hostingEnvironmentProfile.id')) {
                $details[$path] = Get-Field $p $path
            }
        }
        '^microsoft.compute/disks$' {
            $basis = 'Managed disk storage SKU'
            foreach ($path in @('diskSizeGB', 'diskIOPSReadWrite', 'diskMBpsReadWrite', 'diskState', 'encryption.type')) {
                $details[$path] = Get-Field $p $path
            }
            $guidance = 'Plan disk copy/replication, encryption-key access, supported disk SKU and zone constraints; provisioned GB is not used data.'
        }
        '^microsoft.storage/storageaccounts$' {
            $basis = 'Storage redundancy/performance SKU'
            foreach ($path in @('accessTier', 'primaryLocation', 'secondaryLocation', 'isHnsEnabled',
                'isSftpEnabled', 'allowBlobPublicAccess', 'largeFileSharesState', 'networkAcls.defaultAction')) {
                $details[$path] = Get-Field $p $path
            }
            $guidance = 'Inventory blobs/files/queues/tables separately; validate replication, HNS, private endpoints, throughput, data volume and endpoint changes.'
        }
        '^microsoft.sql/(servers/databases|servers/elasticpools|managedinstances)$' {
            $basis = 'SQL service objective / compute tier'
            if (-not $sku) { $sku = Get-Field $p 'currentServiceObjectiveName' }
            foreach ($path in @('elasticPoolId', 'maxSizeBytes', 'currentServiceObjectiveName',
                'zoneRedundant', 'readScale', 'licenseType', 'requestedBackupStorageRedundancy',
                'vCores', 'storageSizeInGB', 'subnetId')) {
                $details[$path] = Get-Field $p $path
            }
            $guidance = 'Validate geo-replication or export/restore path, failover groups, logins, TDE/CMK, pool membership, RPO/RTO and connectivity.'
        }
        '^microsoft.containerservice/managedclusters$' {
            $basis = 'AKS control-plane SKU; node pool VM sizes are in ServiceDetails'
            $details.KubernetesVersion = Get-Field $p 'kubernetesVersion'
            $details.NodeResourceGroup = Get-Field $p 'nodeResourceGroup'
            $details.NetworkPlugin = Get-Field $p 'networkProfile.networkPlugin'
            $nodePools = Get-Field $p 'agentPoolProfiles'
            $details.NodePools = @(foreach ($pool in $nodePools) {
                if ($pool) {
                    [ordered]@{ Name = $pool['name']; VMSize = $pool['vmSize']; Count = $pool['count']
                        Min = $pool['minCount']; Max = $pool['maxCount']; Zones = $pool['availabilityZones']
                        SubnetId = $pool['vnetSubnetID']; Mode = $pool['mode']; OsType = $pool['osType'] }
                }
            })
            $guidance = 'Rebuild cluster; migrate manifests/secrets securely, persistent volumes, ingress, identities, registries and node-pool quotas. Inventory is not Kubernetes workload discovery.'
        }
        '^microsoft.documentdb/databaseaccounts$' {
            $basis = 'Cosmos account capabilities; throughput is usually on child resources'
            foreach ($path in @('databaseAccountOfferType', 'locations', 'enableMultipleWriteLocations',
                'enableAutomaticFailover', 'capabilities', 'consistencyPolicy.defaultConsistencyLevel')) {
                $details[$path] = Get-Field $p $path
            }
            $guidance = 'Review all account regions, failover priorities, write regions, consistency, throughput, private endpoints and residency requirements.'
        }
        '^microsoft.cache/(redis|redisenterprise)$' {
            $basis = 'Redis SKU/family/capacity'
            if (-not $sku) {
                $sku = Get-Field $p 'sku.name'
                $capacity = Get-Field $p 'sku.capacity'
            }
            $details.Family = Get-Field $Resource 'sku.family'
            if (-not $details.Family) { $details.Family = Get-Field $p 'sku.family' }
            foreach ($path in @('shardCount', 'replicasPerMaster', 'minimumTlsVersion', 'subnetId')) {
                $details[$path] = Get-Field $p $path
            }
            $guidance = 'Check target service/SKU and lifecycle; plan replication/export or cache warm-up and endpoint cutover.'
        }
        '^microsoft.dbfor(postgresql|mysql)/flexibleservers$' {
            $basis = 'Database compute SKU/tier'
            foreach ($path in @('version', 'storage.storageSizeGB', 'highAvailability.mode',
                'backup.backupRetentionDays', 'backup.geoRedundantBackup', 'network.delegatedSubnetResourceId')) {
                $details[$path] = Get-Field $p $path
            }
            $guidance = 'Validate replica/restore migration path, engine/version/extensions, HA, delegated subnet, private DNS and data loss window.'
        }
        '^microsoft.app/(containerapps|managedenvironments)$' {
            $basis = 'Container Apps profile; consumption may have no resource SKU'
            foreach ($path in @('environmentId', 'managedEnvironmentId', 'workloadProfileName',
                'workloadProfiles', 'template.scale.minReplicas', 'template.scale.maxReplicas')) {
                $details[$path] = Get-Field $p $path
            }
            $guidance = 'Recreate environment/apps and validate workload profiles, registries, identities, certificates, ingress and stateful dependencies.'
        }
        '^microsoft.keyvault/vaults$' {
            $sku = Get-Field $p 'sku.name'
            $basis = 'Key Vault SKU'
            foreach ($path in @('enableSoftDelete', 'enablePurgeProtection', 'softDeleteRetentionInDays',
                'enableRbacAuthorization')) { $details[$path] = Get-Field $p $path }
            $guidance = 'Plan supported secret/key/certificate migration and CMK dependencies; do not export secret values into inventory.'
        }
        '^microsoft.network/virtualnetworks$' {
            $details.AddressPrefixes = Get-Field $p 'addressSpace.addressPrefixes'
            $details.DnsServers = Get-Field $p 'dhcpOptions.dnsServers'
            $guidance = 'Plan address-space overlap, routing, firewall/NVA, peering, DNS, VPN/ExpressRoute and connectivity cutover.'
        }
        '^microsoft.network/(publicipaddresses|loadbalancers|applicationgateways|azurefirewalls|privateendpoints)$' {
            foreach ($path in @('publicIPAllocationMethod', 'publicIPAddressVersion',
                'sku.name', 'gatewayIPConfigurations', 'subnet.id')) {
                $details[$path] = Get-Field $p $path
            }
            $guidance = 'Validate target network SKU/zones, private-link approvals, DNS, certificates, allowlists and IP/endpoint changes.'
        }
    }
    if (-not $sku) { $sku = 'Not exposed / not applicable' }
    [pscustomobject]@{
        Name = [string]$sku; Tier = $tier; Capacity = $capacity; Basis = $basis
        Source = $source; Details = ConvertTo-CompactJson $details; Guidance = $guidance
    }
}

function New-JsonLineWriter {
    param([string]$Path)
    [IO.StreamWriter]::new($Path, $false, [Text.UTF8Encoding]::new($false))
}

function Write-JsonLine {
    param([IO.StreamWriter]$Writer, $Value)
    $Writer.WriteLine((ConvertTo-CompactJson $Value))
}

function Get-JsonLines {
    param([string]$Path)
    $reader = [IO.StreamReader]::new($Path)
    try {
        while ($null -ne ($line = $reader.ReadLine())) {
            if ($line) { ConvertFrom-Json -InputObject $line -AsHashtable -Depth 100 }
        }
    }
    finally { $reader.Dispose() }
}

function ConvertTo-InventoryRow {
    param($Resource, $Subscription, $ServiceProfile, [string]$SourceRegion, [string]$TargetRegion)
    $p = $Resource['properties']
    $secondary = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($path in @('primaryLocation', 'secondaryLocation')) {
        $location = [string](Get-Field $p $path)
        if ($location) { [void]$secondary.Add($location.Replace(' ', '').ToLowerInvariant()) }
    }
    $locations = Get-Field $p 'locations'
    foreach ($location in $locations) {
        $name = [string](Get-Field $location 'locationName')
        if ($name) { [void]$secondary.Add($name.Replace(' ', '').ToLowerInvariant()) }
    }
    $primaryMatch = ([string]$Resource.location -ieq $SourceRegion)
    $secondaryMatch = $secondary.Contains($SourceRegion)
    $scope = if ($primaryMatch) { 'PrimaryRegion' } elseif ($secondaryMatch) { 'SecondaryRegion' } else { 'GlobalContext' }
    [ordered]@{
        ResourceId = $Resource.id; Name = $Resource.name; ResourceType = $Resource.type
        ResourceGroup = $Resource.resourceGroup; SubscriptionId = $Resource.subscriptionId
        SubscriptionName = $Subscription.name; TenantId = $Subscription.tenantId
        Region = $Resource.location; EffectiveRegion = $Resource.location
        ImpactScope = $scope; AdditionalRegions = ($secondary -join ';')
        TargetRegion = $TargetRegion; Kind = $Resource['kind']
        SkuName = $ServiceProfile.Name; SkuTier = $ServiceProfile.Tier; SkuCapacity = $ServiceProfile.Capacity
        SkuBasis = $ServiceProfile.Basis; SkuSourceResourceId = $ServiceProfile.Source
        Zones = (@($Resource['zones']) -join ';'); ProvisioningState = Get-Field $p 'provisioningState'
        ResourceState = Get-Field $p 'state'; ManagedBy = $Resource['managedBy']
        IdentityType = Get-Field $Resource 'identity.type'
        PrincipalId = Get-Field $Resource 'identity.principalId'
        PublicNetworkAccess = Get-Field $p 'publicNetworkAccess'
        TagsJson = ConvertTo-CompactJson $Resource['tags']
        ServiceDetailsJson = $ServiceProfile.Details; MigrationReview = $ServiceProfile.Guidance
        TargetValidation = 'Not performed: availability, quota and capacity require separate validation'
    }
}

#endregion

#region Streaming Excel export
function Get-ExcelColumnName {
    param([int]$Number)
    $name = ''
    while ($Number -gt 0) {
        $Number--
        $name = [char](65 + ($Number % 26)) + $name
        $Number = [int][Math]::Floor($Number / 26)
    }
    $name
}

function Add-ZipText {
    param([IO.Compression.ZipArchive]$Zip, [string]$Name, [string]$Text)
    $entry = $Zip.CreateEntry($Name)
    $writer = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false))
    try { $writer.Write($Text) } finally { $writer.Dispose() }
}

function Write-ExcelCell {
    param([Xml.XmlWriter]$Writer, [string]$Reference, $Value, [int]$Style = 0)
    $Writer.WriteStartElement('c')
    $Writer.WriteAttributeString('r', $Reference)
    if ($Style -gt 0) { $Writer.WriteAttributeString('s', [string]$Style) }
    if ($Value -is [Collections.IDictionary] -and $Value.Contains('Formula')) {
        $Writer.WriteElementString('f', [string]$Value.Formula)
        $Writer.WriteElementString('v', [string]$Value.Cached)
    }
    elseif ($Value -is [byte] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal]) {
        $Writer.WriteElementString('v', [Convert]::ToString($Value, [Globalization.CultureInfo]::InvariantCulture))
    }
    else {
        $text = [string]$Value
        # Excel cells are limited to 32,767 UTF-16 code units. Full values stay in JSONL.
        if ($text.Length -gt 32767) {
            $length = if ([char]::IsHighSurrogate($text[32719])) { 32719 } else { 32720 }
            $text = $text.Substring(0, $length) + ' [TRUNCATED: see JSONL companion]'
        }
        $text = [regex]::Replace($text, '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]', '')
        $Writer.WriteAttributeString('t', 'inlineStr')
        $Writer.WriteStartElement('is')
        $Writer.WriteStartElement('t')
        $Writer.WriteAttributeString('xml', 'space', 'http://www.w3.org/XML/1998/namespace', 'preserve')
        $Writer.WriteString($text)
        $Writer.WriteEndElement(); $Writer.WriteEndElement()
    }
    $Writer.WriteEndElement()
}

function Export-StreamingWorkbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][object[]]$Sheets,
        [ValidateRange(1, 1048575)][int]$RowsPerSheet = 250000
    )
    $stream = [IO.File]::Open($Path, [IO.FileMode]::CreateNew)
    $zip = [IO.Compression.ZipArchive]::new($stream, [IO.Compression.ZipArchiveMode]::Create)
    $sheetNames = [Collections.Generic.List[string]]::new()
    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $totalRows = 0L
    $datasetNumber = 0
    try {
        foreach ($sheet in $Sheets) {
            $datasetNumber++
            Write-Host "  Excel dataset $datasetNumber/$($Sheets.Count): $($sheet.Name)"
            $columnNames = @(for ($col = 1; $col -le $sheet.Columns.Count; $col++) { Get-ExcelColumnName $col })
            $reader = [IO.StreamReader]::new($sheet.Path)
            try {
                $next = $reader.ReadLine()
                $part = 0
                do {
                    $part++
                    $suffix = if ($part -eq 1) { '' } else { "_$part" }
                    $name = $sheet.Name.Substring(0, [Math]::Min($sheet.Name.Length, 31 - $suffix.Length)) + $suffix
                    $sheetNames.Add($name)
                    $sheetNumber = $sheetNames.Count
                    Write-Host "    Writing worksheet '$name' ..."
                    $lastReportMilliseconds = $timer.ElapsedMilliseconds
                    $entry = $zip.CreateEntry("xl/worksheets/sheet$sheetNumber.xml", [IO.Compression.CompressionLevel]::Fastest)
                    $settings = [Xml.XmlWriterSettings]::new()
                    $settings.Encoding = [Text.UTF8Encoding]::new($false)
                    $settings.CloseOutput = $true
                    $writer = [Xml.XmlWriter]::Create($entry.Open(), $settings)
                    try {
                        $writer.WriteStartDocument()
                        $writer.WriteStartElement('worksheet', $ns)
                        $writer.WriteStartElement('sheetViews')
                        $writer.WriteStartElement('sheetView')
                        $writer.WriteAttributeString('workbookViewId', '0')
                        $writer.WriteStartElement('pane')
                        $writer.WriteAttributeString('ySplit', '1')
                        $writer.WriteAttributeString('topLeftCell', 'A2')
                        $writer.WriteAttributeString('activePane', 'bottomLeft')
                        $writer.WriteAttributeString('state', 'frozen')
                        $writer.WriteEndElement(); $writer.WriteEndElement(); $writer.WriteEndElement()
                        $writer.WriteStartElement('cols')
                        for ($col = 1; $col -le $sheet.Columns.Count; $col++) {
                            $writer.WriteStartElement('col')
                            $writer.WriteAttributeString('min', [string]$col)
                            $writer.WriteAttributeString('max', [string]$col)
                            $width = if ($sheet.Columns[$col - 1] -match 'Json|Review|Notes|Id$|Message') { '55' } else { '25' }
                            $writer.WriteAttributeString('width', $width)
                            $writer.WriteAttributeString('customWidth', '1')
                            $writer.WriteEndElement()
                        }
                        $writer.WriteEndElement()
                        $writer.WriteStartElement('sheetData')
                        $writer.WriteStartElement('row')
                        $writer.WriteAttributeString('r', '1')
                        for ($col = 1; $col -le $sheet.Columns.Count; $col++) {
                            Write-ExcelCell $writer "$($columnNames[$col - 1])1" $sheet.Columns[$col - 1] 1
                        }
                        $writer.WriteEndElement()
                        $rowNumber = 1
                        while ($null -ne $next -and $rowNumber -le $RowsPerSheet) {
                            $data = ConvertFrom-Json -InputObject $next -AsHashtable -Depth 100
                            $rowNumber++
                            $writer.WriteStartElement('row')
                            $writer.WriteAttributeString('r', [string]$rowNumber)
                            for ($col = 1; $col -le $sheet.Columns.Count; $col++) {
                                Write-ExcelCell $writer "$($columnNames[$col - 1])$rowNumber" $data[$sheet.Columns[$col - 1]]
                            }
                            $writer.WriteEndElement()
                            $totalRows++
                            if (($rowNumber - 1) % 1000 -eq 0 -or
                                $timer.ElapsedMilliseconds - $lastReportMilliseconds -ge 5000) {
                                Write-Host ("    Excel '{0}': {1:N0} rows; {2:N0} total rows; elapsed {3:N1}s" -f
                                    $name, ($rowNumber - 1), $totalRows, $timer.Elapsed.TotalSeconds)
                                $lastReportMilliseconds = $timer.ElapsedMilliseconds
                            }
                            $next = $reader.ReadLine()
                        }
                        $writer.WriteEndElement()
                        $writer.WriteStartElement('autoFilter')
                        $writer.WriteAttributeString('ref', "A1:$(Get-ExcelColumnName $sheet.Columns.Count)$rowNumber")
                        $writer.WriteEndElement()
                        $writer.WriteEndElement()
                        $writer.WriteEndDocument()
                    }
                    finally { $writer.Dispose() }
                    Write-Host ("    Finished worksheet '{0}': {1:N0} rows; elapsed {2:N1}s" -f
                        $name, ($rowNumber - 1), $timer.Elapsed.TotalSeconds)
                } while ($null -ne $next)
            }
            finally { $reader.Dispose() }
        }
        Write-Host "  Finalizing Excel metadata ($($sheetNames.Count) worksheets; $totalRows total data rows) ..."
        $overrides = ''
        $rels = ''
        $workbookSheets = ''
        for ($i = 1; $i -le $sheetNames.Count; $i++) {
            $escaped = [Security.SecurityElement]::Escape($sheetNames[$i - 1])
            $workbookSheets += "<sheet name=`"$escaped`" sheetId=`"$i`" r:id=`"rId$i`"/>"
            $rels += "<Relationship Id=`"rId$i`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet`" Target=`"worksheets/sheet$i.xml`"/>"
            $overrides += "<Override PartName=`"/xl/worksheets/sheet$i.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`"/>"
        }
        $stylesId = $sheetNames.Count + 1
        $rels += "<Relationship Id=`"rId$stylesId`" Type=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles`" Target=`"styles.xml`"/>"
        Add-ZipText $zip 'xl/workbook.xml' "<?xml version=`"1.0`" encoding=`"UTF-8`"?><workbook xmlns=`"$ns`" xmlns:r=`"http://schemas.openxmlformats.org/officeDocument/2006/relationships`"><sheets>$workbookSheets</sheets><calcPr calcId=`"191029`" fullCalcOnLoad=`"1`"/></workbook>"
        Add-ZipText $zip 'xl/_rels/workbook.xml.rels' "<?xml version=`"1.0`" encoding=`"UTF-8`"?><Relationships xmlns=`"http://schemas.openxmlformats.org/package/2006/relationships`">$rels</Relationships>"
        Add-ZipText $zip '_rels/.rels' '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        Add-ZipText $zip '[Content_Types].xml' "<?xml version=`"1.0`" encoding=`"UTF-8`"?><Types xmlns=`"http://schemas.openxmlformats.org/package/2006/content-types`"><Default Extension=`"rels`" ContentType=`"application/vnd.openxmlformats-package.relationships+xml`"/><Default Extension=`"xml`" ContentType=`"application/xml`"/><Override PartName=`"/xl/workbook.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml`"/><Override PartName=`"/xl/styles.xml`" ContentType=`"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml`"/>$overrides</Types>"
        Add-ZipText $zip 'xl/styles.xml' @'
<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@
    }
    finally {
        Write-Host '  Closing Excel ZIP archive ...'
        try { $zip.Dispose() } finally { $stream.Dispose(); $timer.Stop() }
    }
    Write-Host ("  Workbook export complete: {0:N0} data rows in {1:N1}s." -f $totalRows, $timer.Elapsed.TotalSeconds)
    return ,$sheetNames.ToArray()
}

#endregion

#region Preflight and inventory orchestration
function Test-InventoryPrerequisites {
    [CmdletBinding()]
    param(
        [string]$SourceRegion, [string]$TargetRegion, [string[]]$SubscriptionId,
        [string]$TenantId, [string]$OutputDirectory, [switch]$Login, [switch]$UseDeviceCode
    )
    $ErrorActionPreference = 'Stop'
    $checks = [Collections.Generic.List[object]]::new()
    function Add-PreflightCheck {
        param([string]$Check, [string]$Details)
        $checks.Add([pscustomobject]@{ Check = $Check; Status = 'Passed'; Details = $Details })
        Write-Host "[PASS] $Check : $Details"
    }
    if ($PSVersionTable.PSVersion -lt [version]'7.2') {
        throw 'PowerShell 7.2+ is required. Open pwsh, not Windows PowerShell 5.1.'
    }
    if (-not [Environment]::Is64BitProcess) { throw '64-bit PowerShell is required for this large-estate collector. Open a 64-bit pwsh terminal.' }
    Add-PreflightCheck 'PowerShell' "$($PSVersionTable.PSVersion), 64-bit"
    foreach ($typeName in @('System.IO.Compression.ZipArchive', 'System.Xml.XmlWriter')) {
        if (-not ($typeName -as [type])) { throw "Required .NET type missing: $typeName. Repair your PowerShell 7 installation." }
    }
    Add-PreflightCheck 'Runtime dependencies' '.NET ZIP/XML available; no Azure CLI, ImportExcel, Excel or Python dependency'
    if ($UseDeviceCode -and -not $Login) { throw 'UseDeviceCode requires Login.' }
    $modules = @(Import-InventoryAzModules)
    Add-PreflightCheck 'Az PowerShell modules' (($modules | ForEach-Object { "$($_.Name) $($_.Version)" }) -join '; ')
    if ($Login) {
        $loginOptions = @{ Scope = 'Process'; ErrorAction = 'Stop'; Verbose = $false }
        if ($TenantId) { $loginOptions.Tenant = $TenantId }
        if ($UseDeviceCode) { $loginOptions.UseDeviceAuthentication = $true }
        $null = Connect-AzAccount @loginOptions
    }
    $originalContext = Get-AzContext -ErrorAction Stop -Verbose:$false
    if ($null -eq $originalContext -or $null -eq $originalContext.Account) {
        throw "No Az PowerShell session. Run Connect-AzAccount -Tenant '<tenant-guid>' or use -Login. Azure CLI login does not authenticate Az PowerShell."
    }
    $effectiveTenantId = if ($TenantId) { $TenantId } else { [string]$originalContext.Tenant.Id }
    if (-not $effectiveTenantId) { throw 'No tenant in the current Az context. Specify -TenantId or sign in with Connect-AzAccount.' }
    try {
        $subscriptions = @(Get-AzSubscription -TenantId $effectiveTenantId -DefaultProfile $originalContext -ErrorAction Stop -Verbose:$false)
        $accounts = @($subscriptions | ForEach-Object {
            @{ id = $_.Id; name = $_.Name; tenantId = $_.TenantId; state = [string]$_.State }
        })
    }
    catch {
        throw "Az subscription discovery failed. Run Connect-AzAccount for the intended tenant and verify subscription Reader access. Original error: $($_.Exception.Message)"
    }
    $accounts = @($accounts | Where-Object { $_.tenantId -eq $effectiveTenantId })
    $requestedIds = @()
    if ($SubscriptionId) { $requestedIds = @($SubscriptionId | Select-Object -Unique) }
    if ($SubscriptionId) {
        $visibleIds = @($accounts | ForEach-Object { $_.id })
        foreach ($requested in $SubscriptionId) {
            if ($requested -notin $visibleIds) {
                throw "Subscription not visible in tenant $effectiveTenantId : $requested. Check tenant, ID and Reader access. Cross-tenant arrays are not supported; run separately for each tenant."
            }
        }
        $accounts = @($accounts | Where-Object { $_.id -in $SubscriptionId })
    }
    if ($accounts.Count -eq 0) { throw 'No subscriptions in the requested scope. Check tenant, Connect-AzAccount and subscription Reader access.' }
    $scopeType = if ($SubscriptionId) { 'Subscriptions' } else { 'Tenant' }
    $scopeLabel = if ($SubscriptionId) { "$($requestedIds.Count) subscriptions in tenant $effectiveTenantId" } else { "Tenant $effectiveTenantId" }
    $nonEnabledAccounts = @($accounts | Where-Object { $_.state -ne 'Enabled' })
    foreach ($account in $nonEnabledAccounts) {
        Write-Warning "Subscription $($account.id) ($($account.name)) is $($account.state), not Enabled. It remains in the ARG query scope; Azure may not expose all of its resources. Review its state with the subscription administrator."
    }
    $stateSummary = if ($nonEnabledAccounts.Count) {
        'Non-Enabled subscriptions retained in query scope: ' +
        (($nonEnabledAccounts | ForEach-Object { "$($_.id) ($($_.state))" }) -join '; ') +
        '. ARG visibility is not proof of complete inventory for these subscriptions.'
    } else { 'All discovered subscriptions in the requested scope are Enabled.' }
    Add-PreflightCheck 'Account scope' "$scopeLabel; $($accounts.Count) visible subscription metadata records. $stateSummary"
    $enabledAccounts = @($accounts | Where-Object { $_.state -eq 'Enabled' })
    if ($enabledAccounts.Count -eq 0) {
        throw 'No Enabled subscription is available in the requested scope to establish the Azure context. Include an Enabled subscription from the same tenant, or ask the subscription administrator to investigate subscription states.'
    }
    try {
        $context = Set-AzContext -Subscription $enabledAccounts[0].id -Tenant $effectiveTenantId -Scope Process -ErrorAction Stop -Verbose:$false
        $initialCount = Get-ArgCount -SubscriptionId $requestedIds -BaseQuery (Get-RegionalResourceQuery $SourceRegion) -DefaultProfile $context
    }
    catch {
        throw "Live ARG preflight failed for $scopeLabel. Check Connect-AzAccount, Reader access, Az compatibility and network/proxy access. Tenant subscription-scope limits may require an explicit smaller subscription array; partial scopes are not enabled. Original error: $($_.Exception.Message)"
    }
    finally { $null = Set-AzContext -Context $originalContext -Scope Process -ErrorAction Stop -Verbose:$false }
    Add-PreflightCheck 'Live Azure Resource Graph access' "$scopeLabel; initial regional Resources count=$initialCount"
    $locations = @(Get-AzLocation -DefaultProfile $context -ErrorAction Stop -Verbose:$false)
    $regionNames = @($locations | ForEach-Object { $_.Location })
    if ($SourceRegion -notin $regionNames) { throw "Unknown source region '$SourceRegion'. Use a canonical region name from Get-AzLocation." }
    if ($TargetRegion -and $TargetRegion -notin $regionNames) { throw "Unknown target region '$TargetRegion'." }
    if ($TargetRegion -and $SourceRegion -eq $TargetRegion) { throw 'SourceRegion and TargetRegion must be different.' }
    Add-PreflightCheck 'Regions' "Source=$SourceRegion; target=$TargetRegion (names only; not target readiness)"
    if (Test-Path -LiteralPath $OutputDirectory) { throw "Output directory already exists; choose a new path: $OutputDirectory" }
    $outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
    $parent = [IO.DirectoryInfo]::new([IO.Path]::GetDirectoryName($outputPath))
    while ($null -ne $parent -and -not $parent.Exists) { $parent = $parent.Parent }
    if ($null -eq $parent) { throw "No existing parent directory found for output: $outputPath" }
    $probePath = Join-Path $parent.FullName ('.inventory-preflight-' + [guid]::NewGuid() + '.tmp')
    $probe = $null
    try {
        $probe = [IO.FileStream]::new($probePath, [IO.FileMode]::CreateNew, [IO.FileAccess]::ReadWrite,
            [IO.FileShare]::None, 4096, [IO.FileOptions]::DeleteOnClose)
        $probe.WriteByte(1)
        $probe.Flush()
    }
    catch { throw "Output parent is not writable: $($parent.FullName). Choose an access-controlled writable path. Original error: $($_.Exception.Message)" }
    finally { if ($null -ne $probe) { $probe.Dispose() } }
    Add-PreflightCheck 'Output path' "Write probe passed in $($parent.FullName); final directory creation is checked during collection"
    [pscustomobject]@{
        Checks = $checks.ToArray(); Accounts = $accounts; OutputDirectory = $outputPath
        InitialCount = $initialCount; Context = $context; TenantId = $effectiveTenantId
        SubscriptionIds = $requestedIds; ScopeType = $scopeType; ScopeLabel = $scopeLabel
        NonEnabledSubscriptions = $nonEnabledAccounts; SubscriptionStateSummary = $stateSummary
    }
}

function Invoke-RegionalInventory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidatePattern('^[a-z0-9]+$')][string]$SourceRegion,
        [ValidatePattern('^[a-z0-9]*$')][string]$TargetRegion = '',
        [ValidateNotNullOrEmpty()][string[]]$SubscriptionId,
        [string]$TenantId,
        [string]$OutputDirectory = (Join-Path $PWD ("RegionalInventory-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))),
        [switch]$Login,
        [switch]$UseDeviceCode,
        [switch]$CheckOnly,
        [ValidateRange(1, 1048575)][int]$RowsPerSheet = 250000
    )
    $ErrorActionPreference = 'Stop'
    $preflight = Test-InventoryPrerequisites -SourceRegion $SourceRegion -TargetRegion $TargetRegion `
        -SubscriptionId $SubscriptionId -TenantId $TenantId -OutputDirectory $OutputDirectory `
        -Login:$Login -UseDeviceCode:$UseDeviceCode
    if ($CheckOnly) { return $preflight.Checks }
    $accounts = @($preflight.Accounts)
    $accountIndex = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($account in $accounts) { $accountIndex[$account.id] = $account }
    $graphOptions = @{ SubscriptionId = $preflight.SubscriptionIds; DefaultProfile = $preflight.Context }
    $auditScope = @{
        QueryScope = $preflight.ScopeLabel; ScopeType = $preflight.ScopeType
        TenantId = $preflight.TenantId; SubscriptionIds = ($preflight.SubscriptionIds -join ';')
    }
    $OutputDirectory = $preflight.OutputDirectory
    $null = New-Item -ItemType Directory -Path $OutputDirectory
    [IO.File]::WriteAllText((Join-Path $OutputDirectory 'Preflight.json'), (ConvertTo-Json -InputObject $preflight.Checks -Depth 10))
    $manifestPath = Join-Path $OutputDirectory 'Run.json'
    $manifest = [ordered]@{
        Status = 'Collecting'; StartedUtc = [DateTime]::UtcNow.ToString('o'); FinishedUtc = ''
        SourceRegion = $SourceRegion; TargetRegion = $TargetRegion; ResourceScope = 'Regional'
        ScopeType = $preflight.ScopeType; TenantId = $preflight.TenantId
        SubscriptionsRequested = $preflight.SubscriptionIds; SubscriptionsDiscovered = @($accounts.id); ScopeFailures = 0
        NonEnabledSubscriptions = $preflight.NonEnabledSubscriptions
        Note = 'Read-only ARG metadata; not a transactionally consistent snapshot or proof of migration readiness.'
    }
    [IO.File]::WriteAllText($manifestPath, (ConvertTo-Json $manifest -Depth 10))
    $paths = @{}
    $writers = @{}
    foreach ($name in @('Inventory', 'GlobalContext', 'Summary',
        'ScopeAudit', 'ResourceGroups', 'Coverage', 'Errors', 'InventoryStaging')) {
        $paths[$name] = Join-Path $OutputDirectory "$name.jsonl"
        $writers[$name] = New-JsonLineWriter $paths[$name]
    }
    $index = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $impacted = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $stats = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $failureCount = 0
    $resourceProjection = 'id, name, type, location, resourceGroup, subscriptionId, kind, sku, tags, zones, managedBy, identity, properties'
    $base = Get-RegionalResourceQuery $SourceRegion
    try {
            Write-Host "Collecting $($preflight.ScopeLabel); regional resources for $SourceRegion plus global/regionless and known secondary-region context. Resource-group metadata covers the full selected scope."
            try {
                $rgAudit = Invoke-ArgPaged @graphOptions -DatasetName ResourceGroups -BaseQuery "ResourceContainers | where type =~ 'microsoft.resources/subscriptions/resourcegroups'" `
                    -Projection 'id, name, location, subscriptionId, tags' -OnPage {
                        param($rows)
                        foreach ($rg in $rows) {
                            Write-JsonLine $writers.ResourceGroups ([ordered]@{
                                ResourceGroupId = $rg.id; ResourceGroup = $rg.name; SubscriptionId = $rg.subscriptionId
                                MetadataLocation = $rg.location; TagsJson = ConvertTo-CompactJson $rg.tags
                            })
                        }
                    }
                Write-JsonLine $writers.ScopeAudit ($auditScope + @{
                    Dataset = 'ResourceContainers'
                    ExpectedBefore = $rgAudit.Expected; Downloaded = $rgAudit.Downloaded; ExpectedAfter = $rgAudit.After
                    Pages = $rgAudit.Pages; Status = 'CountReconciled'
                })
            }
            catch {
                $failureCount++
                Write-Warning "Resource group collection incomplete: $($_.Exception.Message)"
                Write-JsonLine $writers.Errors @{ Stage = 'ResourceContainers'; QueryScope = $preflight.ScopeLabel; Message = $_.Exception.Message }
                Write-JsonLine $writers.ScopeAudit ($auditScope + @{
                    Dataset = 'ResourceContainers'; Status = 'FAILED - partial rows may be present; not a complete inventory'
                })
            }
            try {
                $pagingOptions = $graphOptions.Clone()
                $pagingOptions.InitialCount = $preflight.InitialCount
                $audit = Invoke-ArgPaged @pagingOptions -BaseQuery $base -Projection $resourceProjection -OnPage {
                    param($rows)
                    foreach ($r in $rows) {
                        $id = [string]$r.id
                        if (-not $accountIndex.ContainsKey([string]$r.subscriptionId)) {
                            throw "Resource $id is outside the discovered subscription metadata scope. Verify tenant visibility and rerun."
                        }
                        $index[$id] = [pscustomobject]@{
                            Id = $id; Name = $r.name; Type = $r.type; Location = $r.location
                            SubscriptionId = $r.subscriptionId; ResourceGroup = $r.resourceGroup
                            SkuName = Get-Field $r 'sku.name'; SkuTier = Get-Field $r 'sku.tier'
                            SkuCapacity = Get-Field $r 'sku.capacity'
                        }
                        $serviceProfile = Get-ServiceProfile -Resource $r -Index $index
                        $row = ConvertTo-InventoryRow $r $accountIndex[[string]$r.subscriptionId] $serviceProfile $SourceRegion $TargetRegion
                        if ($row.ImpactScope -ne 'GlobalContext') {
                            [void]$impacted.Add($id)
                            Write-JsonLine $writers.InventoryStaging $row
                        }
                        elseif (-not $r.location -or $r.location -ieq 'global') {
                            Write-JsonLine $writers.InventoryStaging $row
                        }
                    }
                }
                Write-JsonLine $writers.ScopeAudit ($auditScope + @{
                    Dataset = 'Resources'
                    ExpectedBefore = $audit.Expected; Downloaded = $audit.Downloaded; ExpectedAfter = $audit.After
                    Pages = $audit.Pages; Status = 'CountReconciled'
                })
            }
            catch {
                $failureCount++
                Write-Warning "Resource collection incomplete: $($_.Exception.Message)"
                Write-JsonLine $writers.Errors @{ Stage = 'Resources'; QueryScope = $preflight.ScopeLabel; Message = $_.Exception.Message }
                Write-JsonLine $writers.ScopeAudit ($auditScope + @{
                    Dataset = 'Resources'
                    Status = 'FAILED - partial rows may be present; not a complete inventory'
                })
            }
        Write-Host 'Finalizing inventory regions and service-specific SKUs ...'
        # Correlate after every page is indexed: cross-subscription plans and late pages can resolve.
        $writers.InventoryStaging.Dispose()
        Get-JsonLines $paths.InventoryStaging | ForEach-Object {
            $row = $_
            if ($row.ImpactScope -eq 'GlobalContext' -and -not $row.Region) {
                $candidate = [string]$row.ResourceId
                while (($candidate -split '/').Count -ge 11) {
                    $parts = $candidate -split '/'
                    $candidate = $parts[0..($parts.Count - 3)] -join '/'
                    if ($index.ContainsKey($candidate) -and $index[$candidate].Location) {
                        if ($index[$candidate].Location -ieq $SourceRegion) {
                            $row.ImpactScope = 'InheritedParentRegion'
                            $row.EffectiveRegion = $SourceRegion
                            [void]$impacted.Add($row.ResourceId)
                        }
                        break
                    }
                }
            }
            if ($row.ResourceType -imatch '^microsoft.web/sites(/slots)?$' -and $row.SkuSourceResourceId -and $index.ContainsKey($row.SkuSourceResourceId)) {
                $plan = $index[$row.SkuSourceResourceId]
                $row.SkuName = if ($plan.SkuName) { $plan.SkuName } else { 'Not exposed / not applicable' }
                $row.SkuTier = $plan.SkuTier
                $row.SkuCapacity = $plan.SkuCapacity
                $details = ConvertFrom-Json $row.ServiceDetailsJson -AsHashtable
                [void]$details.Remove('PlanResolution')
                $details['PlanLocation'] = $plan.Location
                $row.ServiceDetailsJson = ConvertTo-CompactJson $details
            }
            if ($row.ImpactScope -eq 'GlobalContext') {
                Write-JsonLine $writers.GlobalContext $row
                return
            }
            Write-JsonLine $writers.Inventory $row
            $key = "$($row.SubscriptionId)|$($row.ResourceType)|$($row.ImpactScope)"
            if (-not $stats.ContainsKey($key)) {
                $stats[$key] = [ordered]@{
                    SubscriptionName = $row.SubscriptionName; SubscriptionId = $row.SubscriptionId
                    ResourceType = $row.ResourceType; ImpactScope = $row.ImpactScope; ResourceCount = 0L
                    MissingTags = 0L; SkuUnknown = 0L
                }
            }
            $stats[$key].ResourceCount++
            if (-not $row.TagsJson -or $row.TagsJson -eq '{}') { $stats[$key].MissingTags++ }
            if ($row.SkuName -eq 'Not exposed / not applicable') { $stats[$key].SkuUnknown++ }
        }
        foreach ($key in ($stats.Keys | Sort-Object)) { Write-JsonLine $writers.Summary $stats[$key] }
        if ($stats.Count -gt 0) {
            $total = 0L; $missing = 0L; $unknown = 0L
            foreach ($s in $stats.Values) { $total += $s.ResourceCount; $missing += $s.MissingTags; $unknown += $s.SkuUnknown }
            # Summary groups are snapshot measurements; formula totals remain updateable in Excel.
            Write-JsonLine $writers.Summary ([ordered]@{
                SubscriptionName = 'TOTAL'; ResourceCount = @{ Formula = "SUM(E2:E$($stats.Count + 1))"; Cached = $total }
                MissingTags = @{ Formula = "SUM(F2:F$($stats.Count + 1))"; Cached = $missing }
                SkuUnknown = @{ Formula = "SUM(G2:G$($stats.Count + 1))"; Cached = $unknown }
            })
        }
        $coverage = @(
            @('Run status', $(if ($failureCount) { 'PARTIAL - review Errors and ScopeAudit' } else { 'Requested query scope count-reconciled' })),
            @('Query scope', "$($preflight.ScopeType): $($preflight.ScopeLabel). Default is the current tenant, not all signed-in tenants. Explicit subscriptions are queried together."),
            @('Subscription states', $preflight.SubscriptionStateSummary),
            @('Scope', "$SourceRegion primary location, known storage/Cosmos secondary locations, and regionless children of indexed source-region ARM parents. GlobalContext is separate, not counted as impacted."),
            @('Resource collection', 'Regional filter only: primary source region, known secondary-region fields, and global/regionless context. Resource-group metadata covers the selected scope. No all-regions resource scan.'),
            @('Visibility', 'Only Az PowerShell-visible subscriptions and ARG-indexed resources permitted by RBAC. Hidden subscriptions/resources cannot be proven absent.'),
            @('Pagination', 'One tenant-wide or explicit-subscription-array token chain per dataset; Search-AzGraph -First 1000 -SkipToken. Partial scopes are not allowed. Unique IDs and before/after counts checked; equality cannot detect all concurrent changes.'),
            @('Summary', 'Measured snapshot counts by subscription/type/impact scope. MissingTags means no tags, not missing mandatory business tags. SKU unknown includes not applicable.'),
            @('SKU', 'Service-specific units are not interchangeable. App plans are shared; do not sum app inherited workers. AKS node sizes are distinct from control-plane SKU.'),
            @('Analysis limits', 'No dependency extraction, graph analysis or runtime traffic discovery. App Service plan SKU lookup and parent-region inheritance use only resources included in the regional scan.'),
            @('Global resources', 'Review GlobalContext for DNS, Front Door, Traffic Manager and other global control planes. Unreferenced global resources may still route to the region.'),
            @('Multi-region coverage', 'Explicit storage primary/secondary and Cosmos locations only. Other service replication/backup/DR placement needs service-specific validation.'),
            @('Data not collected', 'Data-plane contents/used data size, secrets, app settings, credentials, RBAC assignments, locks, policies, diagnostic settings, costs, utilization, backup coverage and runtime traffic.'),
            @('Target readiness', 'No target-region compatibility, quota, capacity reservation, pricing, policy or legal-residency check. TargetRegion is planning context, not validation.'),
            @('Migration planning', 'Validate RPO/RTO, downtime, service migration support, names/endpoints, certificates, CMK, identities, private DNS, network routes and external consumers with workload owners.'),
            @('Scale', 'JSONL staging and XLSX streaming; compact resource index remains in RAM. Worksheets split automatically; Summary uses one unsplit sheet. Protect output as confidential metadata.'),
            @('Excel cell limit', 'Text beyond 32767 UTF-16 units is explicitly marked truncated; JSONL companions preserve full exported values. No arbitrary strings are evaluated as formulas.')
        )
        foreach ($item in $coverage) { Write-JsonLine $writers.Coverage ([ordered]@{ Topic = $item[0]; Notes = $item[1] }) }
        foreach ($writer in $writers.Values) { $writer.Dispose() }
        if ($stats.Count + 1 -gt $RowsPerSheet) {
            throw "Summary needs $($stats.Count + 1) data rows; increase RowsPerSheet to keep its formulas on one worksheet."
        }
        $sheets = @(
            @{ Name = 'Summary'; Path = $paths.Summary; Columns = @('SubscriptionName','SubscriptionId','ResourceType','ImpactScope','ResourceCount','MissingTags','SkuUnknown') },
            @{ Name = 'Inventory'; Path = $paths.Inventory; Columns = @('ResourceId','Name','ResourceType','ResourceGroup','SubscriptionId','SubscriptionName','TenantId','Region','EffectiveRegion','ImpactScope','AdditionalRegions','TargetRegion','Kind','SkuName','SkuTier','SkuCapacity','SkuBasis','SkuSourceResourceId','Zones','ProvisioningState','ResourceState','ManagedBy','IdentityType','PrincipalId','PublicNetworkAccess','TagsJson','ServiceDetailsJson','MigrationReview','TargetValidation') },
            @{ Name = 'GlobalContext'; Path = $paths.GlobalContext; Columns = @('ResourceId','Name','ResourceType','ResourceGroup','SubscriptionId','SubscriptionName','Region','SkuName','TagsJson','MigrationReview') },
            @{ Name = 'ResourceGroups'; Path = $paths.ResourceGroups; Columns = @('ResourceGroupId','ResourceGroup','SubscriptionId','MetadataLocation','TagsJson') },
            @{ Name = 'ScopeAudit'; Path = $paths.ScopeAudit; Columns = @('ScopeType','QueryScope','TenantId','SubscriptionIds','Dataset','ExpectedBefore','Downloaded','ExpectedAfter','Pages','Status') },
            @{ Name = 'Coverage'; Path = $paths.Coverage; Columns = @('Topic','Notes') },
            @{ Name = 'Errors'; Path = $paths.Errors; Columns = @('Stage','QueryScope','Message') }
        )
        $workbookPath = Join-Path $OutputDirectory 'RegionalInventory.xlsx'
        Write-Host 'Writing Excel workbook ...'
        $sheetNames = Export-StreamingWorkbook -Path ($workbookPath + '.partial') -Sheets $sheets -RowsPerSheet $RowsPerSheet
        Move-Item -LiteralPath ($workbookPath + '.partial') -Destination $workbookPath
        # Only normalized, selected metadata is staged; arbitrary ARG properties are not persisted.
        [IO.File]::Delete($paths.InventoryStaging)
        $manifest.Status = if ($failureCount) { 'Partial' } else { 'Completed' }
        $manifest.ScopeFailures = $failureCount
        $manifest.FinishedUtc = [DateTime]::UtcNow.ToString('o')
        $manifest.Workbook = $workbookPath
        $manifest.Worksheets = $sheetNames
        $manifest.IndexedResources = $index.Count
        $manifest.ImpactedResources = $impacted.Count
        $manifest.CollectionErrors = @(Get-JsonLines $paths.Errors)
        [IO.File]::WriteAllText($manifestPath, (ConvertTo-Json $manifest -Depth 10))
        if ($failureCount) {
            $failureDetails = ($manifest.CollectionErrors | ForEach-Object { "[$($_.Stage)] $($_.Message)" }) -join [Environment]::NewLine
            throw "PARTIAL inventory saved to $workbookPath. $failureCount datasets failed for the requested query scope; review Errors and ScopeAudit. Original collection error(s):$([Environment]::NewLine)$failureDetails"
        }
        Write-Host "Completed: $workbookPath ($($impacted.Count) impacted resources; $($index.Count) resources indexed)."
        [pscustomobject]@{ Workbook = $workbookPath; Manifest = $manifestPath; ImpactedResources = $impacted.Count; IndexedResources = $index.Count }
    }
    catch {
        if ($manifest.Status -ne 'Partial') {
            $manifest.Status = 'Failed'
            $manifest.FinishedUtc = [DateTime]::UtcNow.ToString('o')
            $manifest.ScopeFailures = $failureCount
            $manifest.FailureMessage = $_.Exception.Message
            [IO.File]::WriteAllText($manifestPath, (ConvertTo-Json $manifest -Depth 10))
        }
        throw
    }
    finally {
        foreach ($writer in $writers.Values) { $writer.Dispose() }
    }
}

#endregion

Invoke-RegionalInventory @PSBoundParameters
