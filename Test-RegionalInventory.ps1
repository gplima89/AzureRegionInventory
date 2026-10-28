#Requires -Version 7.2
[CmdletBinding()]
param([switch]$Scale)
$ErrorActionPreference = 'Stop'
$entryScript = Join-Path $PSScriptRoot 'Export-AzureRegionalInventory.ps1'
$tokens = $null
$parseErrors = $null
$scriptAst = [Management.Automation.Language.Parser]::ParseFile($entryScript, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors) { throw ($parseErrors | Out-String) }
$definitions = @($scriptAst.EndBlock.Statements |
    Where-Object { $_ -is [Management.Automation.Language.FunctionDefinitionAst] } |
    ForEach-Object { $_.Extent.Text })
$testModule = New-Module -Name RegionalInventoryTests -ScriptBlock ([scriptblock]::Create(
    "Set-StrictMode -Version Latest`n" + ($definitions -join "`n")
))
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ("RegionalInventoryTests-" + [guid]::NewGuid())
$null = New-Item -ItemType Directory -Path $testRoot
try {
    & $testModule {
        param($TestRoot, $RunScale, $EntryScript)
        $testState = @{}
        $testState.passed = 0
        function Assert-True {
            param([bool]$Condition, [string]$Message)
            if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
            $testState.passed++
        }
        function Assert-Throws {
            param([scriptblock]$Action, [string]$Pattern)
            $caught = $false
            try { & $Action }
            catch {
                if ($_.Exception.Message -notmatch $Pattern) { throw }
                $caught = $true
            }
            Assert-True $caught "Expected exception: $Pattern"
        }
        function New-TestGraphPage {
            param([object[]]$Data = @(), [string]$SkipToken)
            [pscustomobject]@{ Data = $Data; SkipToken = $SkipToken }
        }
        function Search-AzGraph {
            [CmdletBinding()]
            param([string]$Query, [string[]]$Subscription, [switch]$UseTenantScope, [int]$First, [string]$SkipToken, $DefaultProfile)
            & $testState.graphHandler $PSBoundParameters
        }
        function Start-Sleep { param($Seconds) }
        $subA = '11111111-1111-1111-1111-111111111111'
        $subB = '22222222-2222-2222-2222-222222222222'
        $subC = '33333333-3333-3333-3333-333333333333'
        $prefixA = "/subscriptions/$subA/resourceGroups/workload/providers"
        $prefixB = "/subscriptions/$subB/resourceGroups/shared/providers"
        $testState.pageCount = 2505
        $testState.pageCalls = [Collections.Generic.List[object]]::new()
        $testState.graphHandler = {
            param($a)
            if ($a.Query -match 'summarize') { return New-TestGraphPage @(@{ ResourceCount = $testState.pageCount }) }
            $testState.pageCalls.Add(@{} + $a)
            $first = $a.First
            $start = if ($a.ContainsKey('SkipToken')) { [int]$a.SkipToken } else { 0 }
            $end = [Math]::Min($start + $first, $testState.pageCount)
            $rows = @(for ($i = $start; $i -lt $end; $i++) { @{ id = "/resource/$i" } })
            New-TestGraphPage $rows $(if ($end -lt $testState.pageCount) { [string]$end } else { $null })
        }
        $testState.received = 0L
        $result = Invoke-ArgPaged $subA 'Resources' 'id' { param($rows) $testState.received += $rows.Count }
        Assert-True ($result.Downloaded -eq 2505 -and $result.Pages -eq 3) '2505 resources require three pages.'
        Assert-True ($testState.received -eq 2505) 'Every row reaches callback.'
        Assert-True ($testState.pageCalls[1].SkipToken -eq '1000') 'Second page uses -SkipToken.'
        Assert-True (-not $testState.pageCalls[0].ContainsKey('SkipToken')) 'First page omits -SkipToken.'
        Assert-True ($testState.pageCalls[0].First -eq 1000) 'First is exactly 1000.'
        Assert-True ($testState.pageCalls[0].Subscription[0] -eq $subA) 'Every graph call specifies subscription scope.'
        $testState.countQueries = 0
        $testState.originalPagingHandler = $testState.graphHandler
        $testState.graphHandler = {
            param($a)
            if ($a.Query -match 'summarize') { $testState.countQueries++ }
            & $testState.originalPagingHandler $a
        }
        $result = Invoke-ArgPaged $subA 'Resources' 'id' { param($rows) } -InitialCount 2505
        Assert-True ($result.Downloaded -eq 2505 -and $testState.countQueries -eq 1) 'Preflight count reuse saves initial count call but retains final reconciliation.'
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} -InitialCount 2504 } 'count mismatch'
        $testState.graphHandler = $testState.originalPagingHandler
        $testState.pageCount = 0
        $result = Invoke-ArgPaged $subA 'Resources' 'id' { throw 'No rows expected' }
        Assert-True ($result.Downloaded -eq 0) 'Empty scope succeeds.'
        $testState.pageCount = 2000
        $result = Invoke-ArgPaged $subA 'Resources' 'id' { param($rows) }
        Assert-True ($result.Pages -eq 2) 'Exact multiple of 1000 completes without an extra page.'
        if ($RunScale) {
            $testState.pageCalls.Clear()
            $testState.pageCount = 100001
            $result = Invoke-ArgPaged -BaseQuery 'Resources' -Projection 'id' -OnPage { param($rows) }
            Assert-True ($result.Downloaded -eq 100001 -and $result.Pages -eq 101) '100001 tenant-wide resources in 101 pages.'
            Assert-True (@($testState.pageCalls | Where-Object { -not $_.ContainsKey('UseTenantScope') -or $_.ContainsKey('Subscription') }).Count -eq 0) 'Every large tenant page uses UseTenantScope, not subscription scope.'
        }
        $testState.graphHandler = {
            param($a)
            if ($a.Query -match 'summarize') { return New-TestGraphPage @(@{ ResourceCount = 2 }) }
            [pscustomobject]@{ Unexpected = 'no data' }
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'invalid response'
        $testState.graphHandler = {
            param($a)
            if ($a.Query -match 'summarize') { return New-TestGraphPage @(@{ ResourceCount = 2 }) }
            New-TestGraphPage @(@{ id = 'one' }) 'same'
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'Duplicate'
        $testState.graphHandler = {
            param($a)
            if ($a.Query -match 'summarize') { return New-TestGraphPage @(@{ ResourceCount = 2 }) }
            New-TestGraphPage @(@{ id = 'one' })
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'count mismatch'
        $testState.mockPage = 0
        $testState.graphHandler = {
            param($a)
            if ($a.Query -match 'summarize') { return New-TestGraphPage @(@{ ResourceCount = 3 }) }
            $testState.mockPage++
            New-TestGraphPage @(@{ id = "unique-$($testState.mockPage)" }) 'repeated'
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'repeated continuation'
        $testState.mockCounts = 0
        $testState.graphHandler = {
            param($a)
            if ($a.Query -match 'summarize') {
                $testState.mockCounts++
                return New-TestGraphPage @(@{ ResourceCount = $testState.mockCounts })
            }
            New-TestGraphPage @(@{ id = 'one' })
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'before=1, downloaded=1, after=2'

        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} -MaxConsistencyRetries 1 } 'OnRestart is required'
        $testState.recoveryRows = [Collections.Generic.List[string]]::new()
        foreach ($mode in @('DuplicateOnce','CountDriftOnce','DuplicateAlways','Authorization')) {
            $testState.recoveryMode = $mode
            $testState.recoveryStarts = 0
            $testState.recoveryRestarts = 0
            $testState.recoveryCounts = 0
            $testState.driftServed = $false
            $testState.recoveryRows.Clear()
            $testState.graphHandler = {
                param($a)
                if ($a.Query -match 'summarize') {
                    $testState.recoveryCounts++
                    $count = 3
                    if ($testState.recoveryMode -eq 'CountDriftOnce' -and $testState.recoveryStarts -eq 1 -and -not $testState.driftServed) {
                        $count = 4
                        $testState.driftServed = $true
                    }
                    return New-TestGraphPage @(@{ ResourceCount = $count })
                }
                if (-not $a.ContainsKey('SkipToken')) {
                    $testState.recoveryStarts++
                    if ($testState.recoveryMode -eq 'Authorization') { throw 'AuthorizationFailed recovery fixture' }
                    return New-TestGraphPage @(@{ id = 'one' }, @{ id = 'two' }) 'next'
                }
                if ($testState.recoveryMode -eq 'DuplicateAlways' -or
                    ($testState.recoveryMode -eq 'DuplicateOnce' -and $testState.recoveryStarts -eq 1)) {
                    return New-TestGraphPage @(@{ id = 'two' }, @{ id = 'three' })
                }
                New-TestGraphPage @(@{ id = 'three' })
            }
            $recoveryOptions = @{
                SubscriptionId = @($subA); BaseQuery = 'Resources'; Projection = 'id'; MaxConsistencyRetries = 2
                OnPage = { param($rows) foreach ($row in $rows) { $testState.recoveryRows.Add($row.id) } }
                OnRestart = { param($retry) $testState.recoveryRestarts++; $testState.recoveryRows.Clear() }
            }
            if ($mode -in @('DuplicateOnce','CountDriftOnce')) {
                $recovered = Invoke-ArgPaged @recoveryOptions
                Assert-True ($recovered.Attempts -eq 2 -and $testState.recoveryRestarts -eq 1 -and $recovered.Downloaded -eq 3) "$mode recovers with one clean retry."
                Assert-True (($testState.recoveryRows -join ',') -eq 'one,two,three' -and $testState.recoveryCounts -eq $(if ($mode -eq 'CountDriftOnce') {4} else {3})) "$mode discards prior rows and refreshes the retry count."
            }
            elseif ($mode -eq 'DuplicateAlways') {
                Assert-Throws { Invoke-ArgPaged @recoveryOptions } 'Duplicate resource ID'
                Assert-True ($testState.recoveryStarts -eq 3 -and $testState.recoveryRestarts -eq 2 -and $testState.recoveryRows.Count -eq 2) 'Persistent duplicates stop after three attempts and retain only last-attempt rows.'
                $testState.recoveryStarts = 0
                $testState.recoveryRestarts = 0
                $testState.recoveryRows.Clear()
                $approximate = Invoke-ArgPaged @recoveryOptions -ApproximateOnFinalAttempt -WarningAction SilentlyContinue
                Assert-True ($approximate.Status -eq 'Approximate' -and $approximate.Attempts -eq 3 -and $approximate.DuplicateRows -eq 1 -and $approximate.Downloaded -eq 3) 'ApproximateOnFinalAttempt keeps first occurrences on the third attempt.'
                Assert-True ($testState.recoveryRestarts -eq 2 -and ($testState.recoveryRows -join ',') -eq 'one,two,three') 'Approximate final attempt starts from clean output and passes unique rows only.'
            }
            else {
                Assert-Throws { Invoke-ArgPaged @recoveryOptions } 'AuthorizationFailed'
                Assert-True ($testState.recoveryStarts -eq 1 -and $testState.recoveryRestarts -eq 0) 'Authorization failures do not trigger consistency recovery.'
            }
        }

        $index = [Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        $vm = @{ id = "$prefixA/Microsoft.Compute/virtualMachines/vm"; type = 'Microsoft.Compute/virtualMachines'
            properties = @{ hardwareProfile = @{ vmSize = 'Standard_D4s_v5' } } }
        $serviceProfile = Get-ServiceProfile $vm $index
        Assert-True ($serviceProfile.Name -eq 'Standard_D4s_v5' -and $serviceProfile.Capacity -eq 1) 'VM uses hardwareProfile.vmSize.'
        $redis = @{ id = "$prefixA/Microsoft.Cache/Redis/cache"; type = 'microsoft.cache/redis'
            properties = @{ sku = @{ name = 'Premium'; family = 'P'; capacity = 2 } } }
        $redisSku = Get-ServiceProfile $redis $index
        Assert-True ($redisSku.Name -eq 'Premium' -and $redisSku.Capacity -eq 2) 'Redis supports properties.sku.'
        $vmss = @{ id = "$prefixA/Microsoft.Compute/virtualMachineScaleSets/scale"; type = 'microsoft.compute/virtualmachinescalesets'
            sku = @{ name = 'Standard_D2s_v5'; capacity = 4 }
            properties = @{ virtualMachineProfile = @{ storageProfile = @{ osDisk = @{ osType = 'Linux' } } } } }
        $vmssSku = Get-ServiceProfile $vmss $index
        Assert-True ($vmssSku.Details -match 'Linux' -and $vmssSku.Capacity -eq 4) 'VMSS details use virtualMachineProfile.'
        $pools = @(
            @{ name = 'system'; vmSize = 'Standard_D2s_v5'; count = 3; mode = 'System'; availabilityZones = @('1','2'); vnetSubnetID = "$prefixA/Microsoft.Network/virtualNetworks/vnet/subnets/system" }
            @{ name = 'workers'; vmSize = 'Standard_D4s_v5'; count = 5; mode = 'User'; minCount = 2; maxCount = 8 }
        )
        $aks = @{ id = "$prefixA/Microsoft.ContainerService/managedClusters/cluster"; type = 'microsoft.containerservice/managedclusters'
            properties = @{ kubernetesVersion = '1.33'; agentPoolProfiles = @() } }
        foreach ($poolCount in @(0,1,2)) {
            $aks.properties.agentPoolProfiles = @($pools | Select-Object -First $poolCount)
            $aksProfile = Get-ServiceProfile $aks $index
            $aksDetails = ConvertFrom-Json -InputObject $aksProfile.Details -AsHashtable
            Assert-True ($aksDetails.NodePools.Count -eq $poolCount) "AKS exports exactly $poolCount node pools without nesting the source array."
            if ($poolCount) {
                Assert-True ($aksDetails.NodePools[0].Name -eq 'system' -and $aksDetails.NodePools[0].VMSize -eq 'Standard_D2s_v5' -and $aksDetails.NodePools[0].Count -eq 3) 'AKS node-pool name, VM size and count retain their scalar values.'
                Assert-True ($aksDetails.NodePools[0].Zones.Count -eq 2) 'AKS node-pool zone arrays remain intact.'
            }
            if ($poolCount -eq 2) {
                Assert-True ($aksDetails.NodePools[1].Name -eq 'workers' -and $aksDetails.NodePools[1].Max -eq 8) 'AKS exports every node pool with autoscaling metadata.'
            }
        }
        $aks.properties.Remove('agentPoolProfiles')
        $aksProfile = Get-ServiceProfile $aks $index
        Assert-True ((ConvertFrom-Json -InputObject $aksProfile.Details -AsHashtable).NodePools.Count -eq 0) 'AKS tolerates absent node-pool metadata.'
        $planId = "$prefixB/Microsoft.Web/serverfarms/plan"
        $index[$planId] = [pscustomobject]@{ Id = $planId; SkuName = 'P1v3'; SkuTier = 'PremiumV3'; SkuCapacity = 3; Location = 'eastasia' }
        $app = @{ id = "$prefixA/Microsoft.Web/sites/app"; type = 'microsoft.web/sites'; properties = @{ serverFarmId = $planId } }
        $serviceProfile = Get-ServiceProfile $app $index
        Assert-True ($serviceProfile.Name -eq 'P1v3' -and $serviceProfile.Capacity -eq 3) 'App inherits shared plan SKU across subscriptions.'
        $storage = @{ id = "$prefixA/Microsoft.Storage/storageAccounts/store"; name = 'store'; type = 'microsoft.storage/storageaccounts'
            resourceGroup = 'workload'; subscriptionId = $subA; location = 'eastasia'
            properties = @{ secondaryLocation = 'Southeast Asia' }; sku = @{ name = 'Standard_GRS' } }
        $serviceProfile = Get-ServiceProfile $storage $index
        $row = ConvertTo-InventoryRow $storage @{ name = 'A'; tenantId = 'tenant' } $serviceProfile 'southeastasia' 'eastasia'
        Assert-True ($row.ImpactScope -eq 'SecondaryRegion' -and $row.SkuName -eq 'Standard_GRS') 'Secondary-region storage is impacted.'
        Assert-True ($row.AzureService -eq 'Azure Storage (Blob, Files, Queue, Table)' -and $row.ServiceCategory -eq 'Foundational Services (Ring 0)' -and $row.VmSize -eq '') 'Inventory row carries service category columns.'
        foreach ($case in @(
            @('Standard_NV4as_v4', 'NVv4-series', 2), @('Standard_D4s_v5', 'Dv5 and DSv5-series', 0), @('standard_d4ds_v5', 'Ddv5 and Ddsv5-series', 0)
            @('Standard_D4as_v5', 'Dasv5 and Dadsv5-series', 2), @('Standard_B2ms', 'Bs-series', 0), @('Standard_B2s_v2', 'Bsv2-series', 2)
            @('Standard_DS11-1_v2', 'Dv2 and DSv2-series', 0), @('Standard_E96ias_v4', 'Eav4 and Easv4-series', 2), @('Standard_D8s_v4', 'Dsv4-series', 1)
            @('Standard_F8s_v2', 'Fsv2-series', 1), @('Standard_M128ms', 'M-series', 1), @('Standard_M416s_8_v2', 'Mv2-series', 2)
            @('Standard_NC4as_T4_v3', 'NCasT4_v3-series', 2), @('Standard_ND96amsr_A100_v4', 'NDm_A100_v4-Series', 2), @('Standard_ND96asr_v4', 'NDasrA100_v4-Series', 2)
            @('Standard_L8aos_v4', 'Lsv4, Lasv4, and Laosv4-series', 2), @('Standard_D2_v2_Promo', 'Dv2 and DSv2-series', 0), @('Standard_E64is_v3', 'Ev3 and Esv3-series', 0)
        )) {
            $series = Get-VmSeriesClassification $case[0]
            Assert-True ($series.Listed -and $series.Series -eq $case[1] -and $series.Ring -eq $case[2]) "VM size $($case[0]) maps to $($case[1]) Ring $($case[2])."
        }
        $unlisted = Get-VmSeriesClassification 'Standard_D2ls_v6'
        Assert-True (-not $unlisted.Listed -and $unlisted.Series -eq 'Dlsv6-series' -and $null -eq $unlisted.Ring) 'Unlisted VM series is derived and left unclassified.'
        $vmCategory = Get-ServiceCategory @{ type = 'Microsoft.Compute/virtualMachines'; properties = @{ hardwareProfile = @{ vmSize = 'Standard_NV4as_v4' } } }
        Assert-True ($vmCategory.VmSize -eq 'Standard_NV4as_v4' -and $vmCategory.SkuOrVmSeries -eq 'NVv4-series' -and $vmCategory.ServiceCategory -eq 'Strategic Services (Ring 2)') 'VM ring follows its VM series.'
        $vmCategory = Get-ServiceCategory @{ type = 'microsoft.compute/virtualmachines'; properties = @{} }
        Assert-True ($vmCategory.ServiceCategory -eq 'Not classified' -and $vmCategory.ServiceCategoryBasis -match 'not exposed') 'VM without a size is explicitly not classified.'
        $vmssCategory = Get-ServiceCategory @{ type = 'microsoft.compute/virtualmachinescalesets'; sku = @{ name = 'Standard_NV12s_v3' }; properties = @{} }
        Assert-True ($vmssCategory.ServiceCategory -eq 'Strategic Services (Ring 2)' -and $vmssCategory.ServiceCategoryBasis -match 'most restrictive') 'VMSS uses the most restrictive of service and VM-series rings.'
        $aksCategory = Get-ServiceCategory @{ type = 'microsoft.containerservice/managedclusters'; properties = @{ agentPoolProfiles = $pools } }
        Assert-True ($aksCategory.ServiceCategory -eq 'Foundational Services (Ring 0)' -and $aksCategory.VmSize -eq 'Standard_D2s_v5; Standard_D4s_v5') 'AKS exposes distinct node-pool VM sizes and keeps Ring 0 for Ring 0 series.'
        foreach ($case in @(
            , @(@{ type = 'microsoft.storage/storageaccounts'; kind = 'BlockBlobStorage'; properties = @{} }, 'Azure Storage: Premium Blob Storage', 'Mainstream Services (Ring 1)')
            , @(@{ type = 'microsoft.storage/storageaccounts'; kind = 'StorageV2'; properties = @{ isHnsEnabled = $true } }, 'Azure Data Lake Storage Gen2', 'Foundational Services (Ring 0)')
            , @(@{ type = 'microsoft.compute/disks'; sku = @{ name = 'UltraSSD_LRS' }; properties = @{} }, 'Azure Ultra Disk Storage', 'Strategic Services (Ring 2)')
            , @(@{ type = 'microsoft.network/virtualnetworkgateways'; properties = @{ gatewayType = 'ExpressRoute' } }, 'Azure ExpressRoute', 'Foundational Services (Ring 0)')
            , @(@{ type = 'microsoft.web/sites'; kind = 'functionapp,linux'; properties = @{} }, 'Azure Functions', 'Mainstream Services (Ring 1)')
            , @(@{ type = 'microsoft.web/serverfarms'; sku = @{ tier = 'IsolatedV2' }; properties = @{} }, 'Azure App Service: App Service Environment', 'Mainstream Services (Ring 1)')
            , @(@{ type = 'microsoft.automation/automationaccounts/runbooks'; properties = @{} }, 'Azure Automation', 'Strategic Services (Ring 2)')
            , @(@{ type = 'microsoft.network/networkinterfaces'; properties = @{} }, 'Azure Virtual Network', 'Foundational Services (Ring 0)')
            , @(@{ type = 'microsoft.cognitiveservices/accounts'; kind = 'OpenAI'; properties = @{} }, 'Microsoft Foundry', 'Strategic Services (Ring 2)')
            , @(@{ type = 'microsoft.cache/redisenterprise'; sku = @{ name = 'Balanced_B5' }; properties = @{} }, 'Azure Managed Redis', 'Mainstream Services (Ring 1)')
            , @(@{ type = 'microsoft.compute/virtualmachines/extensions'; properties = @{} }, 'Azure Virtual Machines', 'Not classified')
            , @(@{ type = 'microsoft.cache/redis'; properties = @{} }, '', 'Not classified')
        )) {
            $category = Get-ServiceCategory $case[0]
            Assert-True ($category.AzureService -eq $case[1] -and $category.ServiceCategory -eq $case[2]) "$($case[0].type) classifies as '$($case[1])' / $($case[2])."
        }
        $cosmos = @{ id = "$prefixA/Microsoft.DocumentDB/databaseAccounts/cosmos"; name = 'cosmos'; type = 'microsoft.documentdb/databaseaccounts'
            resourceGroup = 'workload'; subscriptionId = $subA; location = 'eastasia'
            properties = @{ locations = @(@{ locationName = 'East Asia' }, @{ locationName = 'Southeast Asia' }) } }
        foreach ($locationCount in @(1,2)) {
            $cosmos.properties.locations = @(@{ locationName = 'East Asia' })
            if ($locationCount -eq 2) { $cosmos.properties.locations += @{ locationName = 'Southeast Asia' } }
            $cosmosProfile = Get-ServiceProfile $cosmos $index
            $cosmosRow = ConvertTo-InventoryRow $cosmos @{ name = 'A'; tenantId = 'tenant' } $cosmosProfile 'southeastasia' 'eastasia'
            Assert-True (($cosmosRow.AdditionalRegions -split ';').Count -eq $locationCount -and ($cosmosRow.AdditionalRegions -split ';') -contains 'eastasia') "Cosmos iterates $locationCount region entries rather than the array wrapper."
            Assert-True ($cosmosRow.ImpactScope -eq $(if ($locationCount -eq 2) {'SecondaryRegion'} else {'GlobalContext'})) 'Cosmos secondary-region impact follows the individual location entries.'
        }




        Assert-True ((Get-RegionalResourceQuery southeastasia) -match "^Resources \| where location =~ 'southeastasia'") 'Regional queries filter before projection and pagination.'
        Assert-Throws { Get-RegionalResourceQuery "southeastasia' or true" } 'pattern'
        $testState.graphAttempts = 0
        $testState.graphHandler = {
            param($a)
            $testState.graphAttempts++
            if ($testState.graphAttempts -eq 1) { throw 'TooManyRequests 429' }
            New-TestGraphPage @(@{ ok = $true })
        }
        $retry = Invoke-InventoryGraph -Query 'Resources' -SubscriptionId $subA -MaxRetries 1
        Assert-True ($retry.data[0].ok -and $testState.graphAttempts -eq 2) 'Throttle is retried.'
        $testState.graphAttempts = 0
        $testState.graphHandler = { param($a) $testState.graphAttempts++; throw 'AuthorizationFailed' }
        Assert-Throws { Invoke-InventoryGraph -Query 'Resources' -SubscriptionId $subA } 'AuthorizationFailed'
        Assert-True ($testState.graphAttempts -eq 1) 'Authorization errors are not retried.'
        $testState.graphHandler = { param($a) New-TestGraphPage @() 'next-page' }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} -InitialCount 1 } 'empty page with a continuation token'
        if (Get-Module -ListAvailable Az.ResourceGraph) {
            Import-Module Az.ResourceGraph -Verbose:$false
            $nativeCommand = Get-Command Search-AzGraph -CommandType Cmdlet
            $responseType = $nativeCommand.ImplementingType.Assembly.GetExportedTypes() |
                Where-Object Name -eq 'PSResourceGraphResponse`1'
            $testState.nativePage = [Activator]::CreateInstance($responseType.MakeGenericType([psobject]))
            $testState.nativePage.Data = [Collections.Generic.List[psobject]]::new()
            $testState.nativePage.Data.Add([pscustomobject]@{ id = 'typed'; properties = @{ hardwareProfile = @{ vmSize = 'Standard_D2s_v5' } } })
            $testState.nativePage.SkipToken = 'native-token'
            $testState.graphHandler = { param($a) Write-Output -NoEnumerate $testState.nativePage }
            $normalized = Invoke-InventoryGraph -Query 'Resources' -SubscriptionId $subA
            Assert-True ($normalized.skipToken -eq 'native-token') 'Native PSResourceGraphResponse retains continuation metadata.'
            Assert-True ($normalized.data[0].properties.hardwareProfile.vmSize -eq 'Standard_D2s_v5') 'Native nested PSObject metadata normalizes to dictionaries.'
            $testState.nativePage.Data.Clear()
            $testState.nativePage.SkipToken = $null
            $normalized = Invoke-InventoryGraph -Query 'Resources' -SubscriptionId $subA
            Assert-True ($normalized.data.Count -eq 0) 'Native empty response retains its wrapper instead of becoming null.'
        }

        # Full offline orchestration: >1000 resources, late cross-subscription plan and secondary region.
        function New-TestResource {
            param($Id, $Type, $Location, $Properties = @{}, $Sku = $null)
            @{ id = $Id; name = ($Id -split '/')[-1]; type = $Type; location = $Location
                resourceGroup = ($Id -split '/')[4]; subscriptionId = ($Id -split '/')[2]
                kind = ''; sku = $Sku; tags = @{}; zones = @(); managedBy = ''; identity = @{}; properties = $Properties }
        }
        $resourcesA = [Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt 1001; $i++) {
            $resourcesA.Add((New-TestResource "$prefixA/Microsoft.Compute/virtualMachines/vm$i" 'microsoft.compute/virtualmachines' 'southeastasia' @{
                hardwareProfile = @{ vmSize = 'Standard_D2s_v5' }; customData = 'THIS-MUST-NOT-BE-PERSISTED'
            }))
        }
        $appId = "$prefixA/Microsoft.Web/sites/app"
        $resourcesA.Add((New-TestResource $appId 'microsoft.web/sites' 'southeastasia' @{ serverFarmId = $planId }))
        $resourcesA.Add((New-TestResource "$appId/slots/staging" 'microsoft.web/sites/slots' '' @{ serverFarmId = $planId }))
        $resourcesA.Add((New-TestResource "$prefixA/Microsoft.Storage/storageAccounts/store" 'microsoft.storage/storageaccounts' 'eastasia' @{
            secondaryLocation = 'southeastasia'
        } @{ name = 'Standard_GRS' }))
        $resourcesA.Add((New-TestResource "$prefixA/Microsoft.Network/trafficManagerProfiles/global" 'microsoft.network/trafficmanagerprofiles' 'global'))
        $resourcesA.Insert(0, $resourcesA[$resourcesA.Count - 1])
        $resourcesA.RemoveAt($resourcesA.Count - 1)
        $resourcesB = @(
            New-TestResource $planId 'microsoft.web/serverfarms' 'southeastasia' @{} @{ name = 'P1v3'; tier = 'PremiumV3'; capacity = 3 }
            New-TestResource "$prefixB/Microsoft.Example/consumers/consumer" 'microsoft.example/consumers' 'eastasia' @{ upstreamId = $appId }
        )
        $testState.testResources = @{ $subA = @($resourcesA.ToArray()); $subB = $resourcesB }
        $testState.simulateFailure = $false
        $testState.simulateResourceFailure = $false
        $testState.simulateResourceDuplicate = $false
        $testState.simulateResourceDuplicateOnce = $false
        $testState.resourceDuplicateInjected = $false
        $testState.removeResourceOnDuplicate = $false
        $testState.removedResourceId = ''
        $testState.simulateGroupDuplicateOnce = $false
        $testState.groupDuplicateInjected = $false
        $testState.groupsPerSubscription = 1
        $testState.graphCalls = [Collections.Generic.List[object]]::new()
        $testState.graphHandler = {
            param($a)
            $testState.graphCalls.Add(@{} + $a)
            $q = $a.Query
            if ($q -match '^Resources' -and $q -notmatch "^Resources \| where location =~ '([^']+)'") {
                throw 'Resource queries must retain their regional filter, including preflight counts.'
            }
            if ($a.DefaultProfile.Tenant.Id -ne 'tenant-a') { throw 'Context tenant does not match query scope' }
            if ($a.ContainsKey('Subscription') -and $a.ContainsKey('UseTenantScope')) { throw 'Mutually exclusive graph scopes supplied' }
            $ids = if ($a.ContainsKey('UseTenantScope')) { @($subA, $subB) } else { @($a.Subscription) }
            if ($testState.simulateFailure -and $q -match '^ResourceContainers') { throw 'AuthorizationFailed test fixture' }
            if ($testState.simulateResourceFailure -and $q -match '^Resources' -and $a.ContainsKey('SkipToken')) {
                throw 'AuthorizationFailed on continuation test fixture'
            }
            $all = @(foreach ($sub in $ids) {
                if ($q -match '^ResourceContainers') {
                    for ($groupNumber = 0; $groupNumber -lt $testState.groupsPerSubscription; $groupNumber++) {
                        @{ id = "/subscriptions/$sub/resourceGroups/metadata-group-$groupNumber"; name = "metadata-group-$groupNumber"
                            subscriptionId = $sub; location = 'eastasia'; tags = @{ Owner = 'Test' } }
                    }
                } else { $testState.testResources[$sub] }
            })
            if ($q -match "^Resources \| where location =~ '([^']+)'") {
                $region = $Matches[1]
                $all = @($all | Where-Object {
                    $resource = $_
                    $secondaryRegions = @(
                        Get-Field $resource.properties 'primaryLocation'
                        Get-Field $resource.properties 'secondaryLocation'
                    )
                    $resourceLocations = Get-Field $resource.properties 'locations'
                    foreach ($resourceLocation in $resourceLocations) {
                        $secondaryRegions += Get-Field $resourceLocation 'locationName'
                    }
                    $secondaryRegions = @($secondaryRegions | Where-Object { $_ } | ForEach-Object { $_.Replace(' ', '').ToLowerInvariant() })
                    $resource.location -ieq $region -or $resource.location -ieq 'global' -or
                        -not $resource.location -or $region -in $secondaryRegions
                })
            }
            $all = @($all | Sort-Object id)
            if ($q -match 'summarize') { return New-TestGraphPage @(@{ ResourceCount = $all.Count }) }
            $start = if ($a.ContainsKey('SkipToken')) { [int]$a.SkipToken } else { 0 }
            $last = [Math]::Min($start + $a.First, $all.Count)
            $data = @()
            if ($all.Count) { $data = @($all[$start..($last - 1)]) }
            if ($q -match '^Resources' -and $start -gt 0 -and
                ($testState.simulateResourceDuplicate -or ($testState.simulateResourceDuplicateOnce -and -not $testState.resourceDuplicateInjected))) {
                $data[0] = $all[0]
                $testState.resourceDuplicateInjected = $true
                if ($testState.removeResourceOnDuplicate) {
                    $testState.removedResourceId = [string]$all[0].id
                    $resourceSubscription = [string]$all[0].subscriptionId
                    $testState.testResources[$resourceSubscription] = @($testState.testResources[$resourceSubscription] |
                        Where-Object { $_.id -ne $testState.removedResourceId })
                }
            }
            if ($q -match '^ResourceContainers' -and $start -gt 0 -and $testState.simulateGroupDuplicateOnce -and -not $testState.groupDuplicateInjected) {
                $data[0] = $all[0]
                $testState.groupDuplicateInjected = $true
            }
            New-TestGraphPage $data $(if ($last -lt $all.Count) { [string]$last } else { $null })
        }
        $testState.completeHandler = $testState.graphHandler
        $testState.preflightScenario = ''
        $testState.originalContext = [pscustomobject]@{ Account = 'test-account'; Subscription = @{ Id = 'original-subscription' }; Tenant = @{ Id = 'tenant-a' } }
        $testState.currentContext = $testState.originalContext
        $testState.subscriptionCalls = [Collections.Generic.List[object]]::new()
        $testState.loginOptions = $null
        function Import-InventoryAzModules {
            if ($testState.preflightScenario -eq 'MissingModule') { throw 'Required module Az.ResourceGraph is missing. Install-Module Az.ResourceGraph -Scope CurrentUser' }
            if ($testState.preflightScenario -eq 'OldModule') { throw 'Search-AzGraph is missing -SkipToken. Update the owning Az module' }
            foreach ($name in @('Az.Accounts','Az.ResourceGraph','Az.Resources')) {
                [pscustomobject]@{ Name = $name; Version = '1.0-test' }
            }
        }
        function Get-AzContext {
            [CmdletBinding()]param()
            if ($testState.preflightScenario -eq 'LoggedOut') { return $null }
            $testState.currentContext
        }
        function Connect-AzAccount {
            [CmdletBinding()]param($Scope, $Tenant, [switch]$UseDeviceAuthentication)
            $testState.loginOptions = @{} + $PSBoundParameters
        }
        function Get-AzSubscription {
            [CmdletBinding()]param($DefaultProfile, $SubscriptionId, $TenantId)
            $testState.subscriptionCalls.Add(@{} + $PSBoundParameters)
            @(
                [pscustomobject]@{ Id = $subB; Name = 'Shared'; TenantId = 'tenant-a'; State = $(if ($testState.preflightScenario -eq 'AllDisabled') {'Disabled'} elseif ($testState.preflightScenario -in @('Disabled','Warned','PastDue')) {$testState.preflightScenario} else {'Enabled'}) }
                [pscustomobject]@{ Id = $subA; Name = 'Workloads'; TenantId = 'tenant-a'; State = $(if ($testState.preflightScenario -eq 'AllDisabled') {'Disabled'} else {'Enabled'}) }
                [pscustomobject]@{ Id = $subC; Name = 'Other tenant'; TenantId = 'tenant-b'; State = 'Enabled' }
            ) | Where-Object { (-not $SubscriptionId -or $_.Id -eq $SubscriptionId) -and (-not $TenantId -or $_.TenantId -eq $TenantId) }
        }
        function Set-AzContext {
            [CmdletBinding()]param($Subscription, $Tenant, $Context, $Scope)
            if ($Scope -ne 'Process') { throw 'Context changes must be process-scoped.' }
            if ($Context) { $testState.currentContext = $Context }
            else { $testState.currentContext = [pscustomobject]@{ Account = 'test-account'; Subscription = @{ Id = $Subscription }; Tenant = @{ Id = $Tenant } } }
            $testState.currentContext
        }
        function Get-AzLocation {
            [CmdletBinding()]param($DefaultProfile)
            [pscustomobject]@{ Location = 'southeastasia' }
            [pscustomobject]@{ Location = 'eastasia' }
        }
        $checkPath = Join-Path $TestRoot 'preflight-only'
        $checks = @(Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $checkPath -CheckOnly)
        Assert-True ($checks.Count -eq 7) 'CheckOnly validates runtime, Az modules, tenant scope, one live query, regions and path.'
        Assert-True ($testState.currentContext -eq $testState.originalContext) 'Preflight restores original Az context.'
        Assert-True (-not (Test-Path $checkPath)) 'CheckOnly does not create the output directory.'
        Assert-True (@(Get-ChildItem -LiteralPath $TestRoot -Filter '.inventory-preflight-*').Count -eq 0) 'Write probe is cleaned up.'
        $testState.originalContext.Tenant.Id = 'tenant-b'
        try {
            $null = Invoke-RegionalInventory -SourceRegion southeastasia -TenantId 'tenant-a' -OutputDirectory $checkPath -CheckOnly
            Assert-True ($testState.subscriptionCalls[-1].TenantId -eq 'tenant-a' -and $testState.graphCalls[-1].DefaultProfile.Tenant.Id -eq 'tenant-a') 'Explicit TenantId overrides the current tenant for discovery and ARG.'
            Assert-True ($testState.currentContext.Tenant.Id -eq 'tenant-b') 'Explicit tenant preflight restores the original other-tenant context.'
        }
        finally { $testState.originalContext.Tenant.Id = 'tenant-a' }
        $testState.subscriptionCalls.Clear()
        $testState.graphCalls.Clear()
        $null = Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId $subA -OutputDirectory $checkPath -CheckOnly
        Assert-True ($testState.subscriptionCalls.Count -eq 1 -and $testState.subscriptionCalls[0].TenantId -eq 'tenant-a') 'Subscription metadata is discovered once for the current tenant.'
        Assert-True ($testState.graphCalls.Count -eq 1) 'Explicit scope still verifies live ARG access.'
        Assert-True ($testState.graphCalls[0].Subscription[0] -eq $subA -and -not $testState.graphCalls[0].ContainsKey('UseTenantScope')) 'Explicit scope is forwarded without tenant expansion.'
        $testState.graphCalls.Clear()
        $null = Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId @($subA,$subB,$subA) -OutputDirectory $checkPath -CheckOnly
        Assert-True ($testState.graphCalls.Count -eq 1 -and $testState.graphCalls[0].Subscription.Count -eq 2) 'Subscription array is deduplicated and queried together in one preflight count.'
        $null = Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId $subA -OutputDirectory $checkPath -CheckOnly -Login -UseDeviceCode -TenantId 'tenant-a'
        Assert-True ($testState.loginOptions.UseDeviceAuthentication -and $testState.loginOptions.Scope -eq 'Process' -and $testState.loginOptions.Tenant -eq 'tenant-a') 'Login maps to process-scoped Az device authentication.'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $checkPath -CheckOnly -UseDeviceCode } 'requires Login'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $TestRoot -CheckOnly } 'already exists'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion invalidregion -OutputDirectory $checkPath -CheckOnly } 'Unknown source region'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -TargetRegion southeastasia -OutputDirectory $checkPath -CheckOnly } 'must be different'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId 'not-visible' -OutputDirectory $checkPath -CheckOnly } 'not visible'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId @($subA,$subC) -OutputDirectory $checkPath -CheckOnly } 'Cross-tenant arrays'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId @() -OutputDirectory $checkPath -CheckOnly } 'empty'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -TenantId 'another-tenant' -OutputDirectory $checkPath -CheckOnly } 'No subscriptions'
        $testState.preflightScenario = ''
        $testState.graphHandler = {
            param($a)
            if ($testState.preflightScenario -eq 'ExpiredToken') { throw 'AADSTS token expired' }
            & $testState.completeHandler $a
        }
        foreach ($case in @(
            @('OldModule', 'Update the owning Az module'),
            @('MissingModule', 'Install-Module'), @('LoggedOut', 'Connect-AzAccount'),
            @('ExpiredToken', 'Live ARG preflight failed'), @('AllDisabled', 'No Enabled subscription')
        )) {
            $testState.preflightScenario = $case[0]
            Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $checkPath -CheckOnly } $case[1]
            Assert-True (-not (Test-Path $checkPath)) "Preflight failure $($case[0]) stops before output collection."
            Assert-True ($testState.currentContext -eq $testState.originalContext) "Context restored after $($case[0]) failure."
        }
        foreach ($state in @('Disabled','Warned','PastDue')) {
            $testState.preflightScenario = $state
            $testState.graphCalls.Clear()
            $stateWarnings = @()
            $statePreflight = Test-InventoryPrerequisites -SourceRegion southeastasia -OutputDirectory $checkPath -WarningVariable stateWarnings
            Assert-True ($statePreflight.Accounts.Count -eq 2 -and $statePreflight.NonEnabledSubscriptions.Count -eq 1 -and $statePreflight.NonEnabledSubscriptions[0].state -eq $state) "State $state is recorded without excluding its subscription."
            Assert-True ($statePreflight.Context.Subscription.Id -eq $subA -and $testState.currentContext -eq $testState.originalContext) "State $state uses an Enabled context and restores the original context."
            Assert-True ($testState.graphCalls.Count -eq 1 -and $testState.graphCalls[0].UseTenantScope -and -not $testState.graphCalls[0].ContainsKey('Subscription')) "State $state preserves tenant-wide ARG scope."
            Assert-True (@($stateWarnings | Where-Object { "$_" -match $subB -and "$_" -match $state }).Count -eq 1) "State $state emits an actionable subscription warning."
            Assert-True ($statePreflight.Checks[3].Details -match $subB) "State $state is recorded in the saved preflight checks."
        }
        $testState.preflightScenario = 'Disabled'
        $stateArray = Test-InventoryPrerequisites -SourceRegion southeastasia -SubscriptionId @($subB,$subA) -OutputDirectory $checkPath
        Assert-True ($stateArray.SubscriptionIds.Count -eq 2 -and $testState.graphCalls[-1].Subscription -contains $subB -and -not $testState.graphCalls[-1].ContainsKey('UseTenantScope')) 'Explicit arrays retain non-Enabled subscriptions without tenant expansion.'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId $subB -OutputDirectory $checkPath -CheckOnly } 'No Enabled subscription'
        $testState.preflightScenario = ''
        $testState.graphHandler = $testState.completeHandler
        $out = Join-Path $TestRoot 'complete'
        $testState.graphCalls.Clear()
        $run = Invoke-RegionalInventory -SourceRegion southeastasia -TargetRegion eastasia -OutputDirectory $out -RowsPerSheet 1000 -InformationVariable runProgress
        $runMessages = ($runProgress | ForEach-Object { [string]$_.MessageData }) -join "`n"
        Assert-True ($runMessages -match 'Retrieved ResourceGroups page' -and $runMessages -match 'Retrieved Resources page') 'Collection progress distinguishes resource groups from resources.'
        Assert-True ($testState.graphCalls.Count -eq 7) 'Two-subscription tenant run uses only seven graph count/page calls.'
        Assert-True (@($testState.graphCalls | Where-Object { -not $_.ContainsKey('UseTenantScope') -or $_.ContainsKey('Subscription') }).Count -eq 0) 'Every default graph query uses tenant scope.'
        Assert-True (@($testState.graphCalls | Where-Object { $_.Query -notmatch 'summarize' -and $_.Query -match '^Resources' }).Count -eq 2) 'One resource token chain covers resources across two subscriptions.'
        Assert-True ($run.ImpactedResources -eq 1005 -and $run.IndexedResources -eq 1006) 'Regional run includes primary, secondary, inherited and global context, but excludes unrelated-region resources.'
        $inventory = @(Get-JsonLines (Join-Path $out 'Inventory.jsonl'))
        $appRow = @($inventory | Where-Object { $_.ResourceId -eq $appId })[0]
        Assert-True ($appRow.SkuName -eq 'P1v3') 'Late plan resolution updates final app inventory.'
        $vmRow = @($inventory | Where-Object { $_.ResourceType -eq 'microsoft.compute/virtualmachines' })[0]
        Assert-True ($vmRow.VmSize -eq 'Standard_D2s_v5' -and $vmRow.SkuOrVmSeries -eq 'Dv5 and DSv5-series' -and $vmRow.ServiceCategory -eq 'Foundational Services (Ring 0)') 'Final inventory carries VM size, series and service category.'
        $categoryRows = @(Get-JsonLines (Join-Path $out 'ResourceCategory.jsonl'))
        $vmCategoryRow = @($categoryRows | Where-Object { $_['ResourceType'] -eq 'microsoft.compute/virtualmachines' })
        Assert-True ($vmCategoryRow.Count -eq 1 -and $vmCategoryRow[0].ResourceCount -eq 1001 -and $vmCategoryRow[0].VmSizes -eq 'Standard_D2s_v5') 'Resource Category aggregates VM type, series and sizes.'
        Assert-True ($categoryRows[-1].ServiceCategory -eq 'TOTAL' -and $categoryRows[-1].ResourceCount.Cached -eq $inventory.Count) 'Resource Category total matches inventory rows.'
        Assert-True ($categoryRows[0].ServiceCategory -eq 'Foundational Services (Ring 0)' -and @($categoryRows | Where-Object { $_['ResourceType'] -eq 'microsoft.web/serverfarms' -and $_.ServiceCategory -eq 'Mainstream Services (Ring 1)' }).Count -eq 1) 'Resource Category is ordered by ring and classifies App Service plans.'
        Assert-True ($appRow.SubscriptionId -eq $subA -and $appRow.SubscriptionName -eq 'Workloads') 'Mixed-scope rows use their own subscription metadata, not the query context default subscription.'
        Assert-True (@(Get-JsonLines (Join-Path $out 'ResourceGroups.jsonl')).Count -eq 2) 'Resource-group dataset collected independently of workload region.'
        $slotRow = @($inventory | Where-Object { $_.ResourceType -eq 'microsoft.web/sites/slots' })[0]
        Assert-True ($slotRow.ImpactScope -eq 'InheritedParentRegion' -and $slotRow.EffectiveRegion -eq 'southeastasia') 'Regionless child inherits indexed parent region.'
        foreach ($removedFile in @('Dependencies.jsonl','DependencyImpact.jsonl','ReferenceStaging.jsonl')) {
            Assert-True (-not (Test-Path (Join-Path $out $removedFile))) "Removed dependency output is not generated: $removedFile."
        }
        Assert-True (@($testState.graphCalls | Where-Object { $_.Query -match '^Resources' -and $_.Query -notmatch '^Resources \| where' }).Count -eq 0) 'Every resource count and page query is region-filtered.'
        Assert-True (-not (Test-Path (Join-Path $out 'InventoryStaging.jsonl'))) 'Staging cleaned after success.'
        $manifest = Get-Content (Join-Path $out 'Run.json') -Raw | ConvertFrom-Json
        Assert-True ($manifest.Status -eq 'Completed') 'Successful manifest.'
        Assert-True ($manifest.ResourceScope -eq 'Regional' -and 'DependencyScope' -notin $manifest.PSObject.Properties.Name) 'Manifest declares regional collection without a dependency-analysis mode.'
        Assert-True ($manifest.ScopeType -eq 'Tenant' -and $manifest.TenantId -eq 'tenant-a' -and $manifest.SubscriptionsRequested.Count -eq 0) 'Default manifest records tenant scope without an implicit explicit subscription list.'
        $audits = @(Get-JsonLines (Join-Path $out 'ScopeAudit.jsonl'))
        Assert-True ($audits.Count -eq 2 -and @($audits | Where-Object { $_.ScopeType -ne 'Tenant' -or $_.Status -ne 'CountReconciled' }).Count -eq 0) 'Audit has one aggregate row per dataset rather than per subscription.'
        Assert-True (@($audits | Where-Object { $_.Dataset -eq 'Resources' -and $_.Downloaded -eq 1006 -and $_.Pages -eq 2 }).Count -eq 1) 'Aggregate regional resource audit has exact count and page total.'
        Assert-True (Test-Path (Join-Path $out 'Preflight.json')) 'Successful preflight checks saved alongside inventory.'
        $arrayOut = Join-Path $TestRoot 'array-scope'
        $testState.graphCalls.Clear()
        $arrayRun = Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId @($subA,$subB) -OutputDirectory $arrayOut
        Assert-True ($arrayRun.IndexedResources -eq 1006 -and $testState.graphCalls.Count -eq 7) 'Explicit array collects regional resources from both subscriptions through one aggregate token chain.'
        Assert-True (@($testState.graphCalls | Where-Object { $_.ContainsKey('UseTenantScope') -or $_.Subscription.Count -ne 2 }).Count -eq 0) 'Every explicit-array query retains the same full array without tenant expansion.'
        $arrayManifest = Get-Content (Join-Path $arrayOut 'Run.json') -Raw | ConvertFrom-Json
        Assert-True ($arrayManifest.ScopeType -eq 'Subscriptions' -and $arrayManifest.SubscriptionsRequested.Count -eq 2) 'Explicit-array scope is persisted accurately.'
        $testState.preflightScenario = 'Disabled'
        $stateOut = Join-Path $TestRoot 'disabled-subscription'
        $stateRun = Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $stateOut
        $stateManifest = Get-Content (Join-Path $stateOut 'Run.json') -Raw | ConvertFrom-Json
        Assert-True ($stateManifest.Status -eq 'Completed' -and $stateRun.IndexedResources -eq 1006 -and $stateManifest.SubscriptionsDiscovered -contains $subB) 'A disabled subscription does not stop collection or remove its regional resources.'
        Assert-True ($stateManifest.NonEnabledSubscriptions.Count -eq 1 -and $stateManifest.NonEnabledSubscriptions[0].id -eq $subB -and $stateManifest.NonEnabledSubscriptions[0].state -eq 'Disabled') 'Run manifest persists the non-Enabled subscription and state.'
        $stateCoverage = @(Get-JsonLines (Join-Path $stateOut 'Coverage.jsonl'))
        Assert-True (@($stateCoverage | Where-Object { $_.Topic -eq 'Subscription states' -and $_.Notes -match $subB -and $_.Notes -match 'not proof of complete inventory' }).Count -eq 1) 'Workbook coverage records the subscription-state visibility caveat.'
        $testState.preflightScenario = ''
        $originalResources = $testState.testResources[$subA]
        try {
            $aksId = "$prefixA/Microsoft.ContainerService/managedClusters/cluster"
            $testState.testResources[$subA] += @(
                New-TestResource $aksId 'microsoft.containerservice/managedclusters' 'southeastasia' @{ agentPoolProfiles = $pools }
                New-TestResource $cosmos.id 'microsoft.documentdb/databaseaccounts' 'eastasia' $cosmos.properties
            )
            $arrayMetadataOut = Join-Path $TestRoot 'array-metadata'
            $arrayMetadataRun = Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $arrayMetadataOut
            $arrayMetadataManifest = Get-Content (Join-Path $arrayMetadataOut 'Run.json') -Raw | ConvertFrom-Json
            $arrayMetadataInventory = @(Get-JsonLines (Join-Path $arrayMetadataOut 'Inventory.jsonl'))
            Assert-True ($arrayMetadataManifest.Status -eq 'Completed' -and $arrayMetadataRun.IndexedResources -eq 1008 -and $arrayMetadataRun.ImpactedResources -eq 1007) 'AKS and Cosmos array metadata complete a full regional paged run without partial inventory.'
            $aksRow = @($arrayMetadataInventory | Where-Object { $_.ResourceId -eq $aksId })[0]
            Assert-True ((ConvertFrom-Json -InputObject $aksRow.ServiceDetailsJson -AsHashtable).NodePools.Count -eq 2) 'Final inventory retains both AKS node pools.'
            Assert-True (@($arrayMetadataInventory | Where-Object { $_.ResourceId -eq $cosmos.id -and $_.ImpactScope -eq 'SecondaryRegion' }).Count -eq 1) 'Final inventory includes the Cosmos account through its secondary region.'
        }
        finally { $testState.testResources[$subA] = $originalResources }

        $zip = [IO.Compression.ZipFile]::OpenRead($run.Workbook)
        try {
            foreach ($entry in $zip.Entries) {
                $reader = [IO.StreamReader]::new($entry.Open())
                try {
                    $text = $reader.ReadToEnd()
                    Assert-True ($text -notmatch 'THIS-MUST-NOT-BE-PERSISTED') 'No arbitrary VM properties in workbook.'
                    $xml = [xml]$text
                    Assert-True ($null -ne $xml.DocumentElement) "Valid XML: $($entry.FullName)"
                }
                finally { $reader.Dispose() }
            }
            $entry = $zip.GetEntry('xl/workbook.xml')
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $wb = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
            Assert-True ('Inventory_2' -in @($wb.workbook.sheets.sheet.name)) 'Inventory sheets split at requested threshold.'
            Assert-True (@($wb.workbook.sheets.sheet.name).Count -eq 9 -and 'Resource Category' -in @($wb.workbook.sheets.sheet.name) -and 'Dependencies' -notin @($wb.workbook.sheets.sheet.name) -and 'DependencyImpact' -notin @($wb.workbook.sheets.sheet.name)) 'Workbook contains the eight inventory/report sheets plus the split inventory remainder.'
            $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $summary = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
            $ns = [Xml.XmlNamespaceManager]::new($summary.NameTable)
            $ns.AddNamespace('s', $summary.DocumentElement.NamespaceURI)
            $totals = $summary.SelectNodes('//s:c[s:f]', $ns)
            Assert-True ($totals.Count -eq 3 -and $totals[0].v -eq '1005') 'Summary formula totals have correct cached values.'
        }
        finally { $zip.Dispose() }
        if (Get-Module -ListAvailable ImportExcel) {
            Import-Module ImportExcel
            $package = Open-ExcelPackage -Path $run.Workbook
            try {
                Assert-True ($package.Workbook.Worksheets['Inventory'].Dimension.Rows -eq 1001) 'Independent EPPlus reader opens first inventory worksheet.'
                Assert-True ($package.Workbook.Worksheets['Inventory_2'].Dimension.Rows -eq 6) 'Independent reader sees exact split remainder.'
                Assert-True ($package.Workbook.Worksheets['Summary'].Cells['E7'].Value -eq 1005) 'Independent reader sees cached total.'
                [OfficeOpenXml.CalculationExtension]::Calculate($package.Workbook)
                Assert-True ($package.Workbook.Worksheets['Summary'].Cells['E7'].Value -eq 1005) 'Independent engine recalculates formula total.'
                $categorySheet = $package.Workbook.Worksheets['Resource Category']
                Assert-True ($categorySheet.Cells[$categorySheet.Dimension.Rows, 6].Value -eq 1005) 'Independent engine recalculates Resource Category total.'
            }
            finally { $package.Dispose() }
        }

        $longPath = Join-Path $TestRoot 'strings.jsonl'
        $w = New-JsonLineWriter $longPath
        Write-JsonLine $w @{ Value = '=HYPERLINK("http://invalid","not executed")' }
        Write-JsonLine $w @{ Value = ('x' * 40000) }
        Write-JsonLine $w @{ Value = (('x' * 32719) + [char]::ConvertFromUtf32(0x1F600) + ('x' * 100)) }
        $w.Dispose()
        $stringXlsx = Join-Path $TestRoot 'strings.xlsx'
        $null = Export-StreamingWorkbook $stringXlsx @(@{ Name = 'Strings'; Path = $longPath; Columns = @('Value') })
        $zip = [IO.Compression.ZipFile]::OpenRead($stringXlsx)
        try {
            $r = [IO.StreamReader]::new($zip.GetEntry('xl/worksheets/sheet1.xml').Open())
            try { $x = [xml]$r.ReadToEnd() } finally { $r.Dispose() }
            Assert-True ($x.worksheet.sheetData.row[1].c.t -eq 'inlineStr') 'Untrusted strings cannot become formulas.'
            Assert-True ($x.worksheet.sheetData.row[2].c.is.t.InnerText.Length -le 32767) 'Excel text limit enforced.'
            Assert-True ($x.worksheet.sheetData.row[2].c.is.t.InnerText -match 'TRUNCATED') 'Truncation explicit.'
            Assert-True ($x.worksheet.sheetData.row[3].c.is.t.InnerText.Length -le 32767) 'Unicode truncation preserves surrogate pairs.'
        }
        finally { $zip.Dispose() }
        $progressPath = Join-Path $TestRoot 'progress.jsonl'
        $progressWriter = New-JsonLineWriter $progressPath
        try {
            for ($i = 0; $i -lt 1001; $i++) { Write-JsonLine $progressWriter @{ Value = $i } }
        }
        finally { $progressWriter.Dispose() }
        $emptyPath = Join-Path $TestRoot 'empty.jsonl'
        [IO.File]::WriteAllText($emptyPath, '')
        $progressSheets = Export-StreamingWorkbook -Path (Join-Path $TestRoot 'progress.xlsx') -RowsPerSheet 1000 -Sheets @(
            @{ Name = 'Progress'; Path = $progressPath; Columns = @('Value') }
            @{ Name = 'Empty'; Path = $emptyPath; Columns = @('Value') }
        ) -InformationVariable exportProgress
        $exportMessages = ($exportProgress | ForEach-Object { [string]$_.MessageData }) -join "`n"
        Assert-True ($progressSheets.Count -eq 3 -and $progressSheets[1] -eq 'Progress_2') 'Export progress does not contaminate the returned worksheet-name array.'
        Assert-True ($exportMessages -match 'Excel dataset 1/2: Progress' -and $exportMessages -match 'Excel dataset 2/2: Empty') 'Export announces dataset position and name.'
        Assert-True ($exportMessages.Contains("Excel 'Progress': $((1000).ToString('N0')) rows;") -and $exportMessages -match 'elapsed [0-9]') 'Export reports row activity at the 1000-row checkpoint with elapsed time.'
        Assert-True ($exportMessages -match "Finished worksheet 'Progress_2': 1 rows" -and $exportMessages -match "Finished worksheet 'Empty': 0 rows") 'Export reports split remainders and empty worksheets correctly.'
        Assert-True ($exportMessages -match 'Finalizing Excel metadata' -and $exportMessages -match 'Closing Excel ZIP archive' -and $exportMessages.Contains("Workbook export complete: $((1001).ToString('N0')) data rows")) 'Export reports finalization, archive closing and exact total data rows.'
        $invalidPath = Join-Path $TestRoot 'invalid.jsonl'
        [IO.File]::WriteAllText($invalidPath, '{invalid')
        $invalidWorkbook = Join-Path $TestRoot 'invalid.xlsx.partial'
        $invalidFailure = ''
        try {
            $null = Export-StreamingWorkbook -Path $invalidWorkbook -Sheets @(
                @{ Name = 'Invalid'; Path = $invalidPath; Columns = @('Value') }
            ) -InformationVariable failedExportProgress
        }
        catch { $invalidFailure = $_.Exception.Message }
        $failedMessages = ($failedExportProgress | ForEach-Object { [string]$_.MessageData }) -join "`n"
        Assert-True ($invalidFailure -match 'JSON' -and $failedMessages -notmatch 'Workbook export complete') 'Failed export surfaces its error without announcing completion.'
        $unlocked = [IO.File]::Open($invalidWorkbook, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
        $unlocked.Dispose()
        Assert-True ($failedMessages -match 'Closing Excel ZIP archive') 'Failed export closes the archive and releases its file handle.'
        $testState.simulateFailure = $true
        $partialOut = Join-Path $TestRoot 'partial'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $partialOut } '(?s)PARTIAL inventory saved.*\[ResourceContainers\] AuthorizationFailed test fixture'
        $partial = Get-Content (Join-Path $partialOut 'Run.json') -Raw | ConvertFrom-Json
        Assert-True ($partial.Status -eq 'Partial' -and (Test-Path (Join-Path $partialOut 'RegionalInventory.xlsx'))) 'Partial workbook preserved with explicit failure status.'
        Assert-True ($partial.ScopeFailures -eq 1 -and $partial.IndexedResources -eq 1006) 'A failed group dataset does not prevent regional resource inventory collection.'
        Assert-True ($partial.CollectionErrors.Count -eq 1 -and $partial.CollectionErrors[0].Stage -eq 'ResourceContainers') 'Partial manifest persists the original dataset error.'
        $testState.simulateFailure = $false
        $testState.simulateResourceFailure = $true
        $resourcePartialOut = Join-Path $TestRoot 'resource-partial'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $resourcePartialOut } 'PARTIAL inventory saved'
        $resourcePartial = Get-Content (Join-Path $resourcePartialOut 'Run.json') -Raw | ConvertFrom-Json
        Assert-True ($resourcePartial.Status -eq 'Partial' -and $resourcePartial.IndexedResources -eq 1000) 'Continuation failure preserves partial rows without claiming success.'
        $testState.simulateResourceFailure = $false
        $testState.simulateResourceDuplicate = $true
        $duplicateOut = Join-Path $TestRoot 'duplicate-approximate'
        $duplicateRun = Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $duplicateOut -WarningVariable duplicateWarnings -WarningAction SilentlyContinue
        $duplicateManifest = Get-Content (Join-Path $duplicateOut 'Run.json') -Raw | ConvertFrom-Json
        Assert-True ($duplicateRun.Status -eq 'Approximate' -and $duplicateManifest.Status -eq 'Approximate' -and $duplicateManifest.ScopeFailures -eq 0) 'Persistent duplicates after three attempts complete as Approximate instead of Partial.'
        Assert-True ($duplicateManifest.ApproximateDatasets.Count -eq 1 -and $duplicateManifest.ApproximateDatasets[0] -match '^Resources: 1 duplicate') 'Manifest records which dataset is approximate.'
        Assert-True ((@($duplicateWarnings) -join "`n") -match 'Inventory is APPROXIMATE') 'A final approximate-inventory warning is printed.'
        $duplicateAudit = @(Get-JsonLines (Join-Path $duplicateOut 'ScopeAudit.jsonl'))
        Assert-True ($duplicateManifest.PagingConsistencyRetries -eq 2 -and $duplicateManifest.IndexedResources -eq 1005 -and
            @($duplicateAudit | Where-Object { $_.Status -eq 'RETRY - attempt discarded' }).Count -eq 2 -and
            @($duplicateAudit | Where-Object { $_.Dataset -eq 'Resources' -and $_.Attempt -eq 3 -and $_.Status -eq 'Approximate' -and $_.DuplicateRows -eq 1 }).Count -eq 1) 'Exhausted strict retries are audited and the final attempt keeps first occurrences only.'
        $duplicateCoverage = @(Get-JsonLines (Join-Path $duplicateOut 'Coverage.jsonl'))
        Assert-True (@($duplicateCoverage | Where-Object { $_.Topic -eq 'Run status' -and $_.Notes -match '^APPROXIMATE' }).Count -eq 1) 'Workbook coverage labels the run approximate.'
        $testState.simulateResourceDuplicate = $false
        $originalRecoveryResources = $testState.testResources[$subA]
        $testState.simulateResourceDuplicateOnce = $true
        $testState.resourceDuplicateInjected = $false
        $testState.removeResourceOnDuplicate = $true
        try {
            $recoveredOut = Join-Path $TestRoot 'recovered-resources'
            $recoveredRun = Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $recoveredOut
            $recoveredManifest = Get-Content (Join-Path $recoveredOut 'Run.json') -Raw | ConvertFrom-Json
            $recoveredRows = @(Get-JsonLines (Join-Path $recoveredOut 'Inventory.jsonl'))
            $recoveredAudit = @(Get-JsonLines (Join-Path $recoveredOut 'ScopeAudit.jsonl'))
            Assert-True ($recoveredManifest.Status -eq 'Completed' -and $recoveredManifest.PagingConsistencyRetries -eq 1 -and
                $recoveredRun.IndexedResources -eq 1005 -and $recoveredRun.ImpactedResources -eq 1004) 'A transient duplicate recovers with fresh counts and no stale index/impact entries after a resource disappears.'
            Assert-True ($recoveredRows.Count -eq 1004 -and @($recoveredRows.ResourceId | Sort-Object -Unique).Count -eq 1004 -and
                $testState.removedResourceId -notin $recoveredRows.ResourceId) 'Resource retry truncates prior staging and exports each current impacted resource once.'
            Assert-True (@($recoveredAudit | Where-Object { $_.Dataset -eq 'Resources' -and $_.Attempt -eq 2 -and $_.ExpectedBefore -eq 1005 -and $_.Status -eq 'CountReconciled' }).Count -eq 1) 'Recovered resource audit records the successful attempt and refreshed expected count.'
        }
        finally {
            $testState.testResources[$subA] = $originalRecoveryResources
            $testState.simulateResourceDuplicateOnce = $false
            $testState.removeResourceOnDuplicate = $false
        }
        $testState.groupsPerSubscription = 1001
        $testState.simulateGroupDuplicateOnce = $true
        try {
            $groupRecoveryOut = Join-Path $TestRoot 'recovered-groups'
            $null = Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId $subA -OutputDirectory $groupRecoveryOut
            $groupRecoveryManifest = Get-Content (Join-Path $groupRecoveryOut 'Run.json') -Raw | ConvertFrom-Json
            $groupRecoveryRows = @(Get-JsonLines (Join-Path $groupRecoveryOut 'ResourceGroups.jsonl'))
            Assert-True ($groupRecoveryManifest.Status -eq 'Completed' -and $groupRecoveryManifest.PagingConsistencyRetries -eq 1) 'Resource-group paging also recovers from a transient duplicate.'
            Assert-True ($groupRecoveryRows.Count -eq 1001 -and @($groupRecoveryRows.ResourceGroupId | Sort-Object -Unique).Count -eq 1001) 'Resource-group retry truncates previously written pages before restarting.'
        }
        finally {
            $testState.groupsPerSubscription = 1
            $testState.simulateGroupDuplicateOnce = $false
        }
        $testState.filteredCalls = [Collections.Generic.List[object]]::new()
        $testState.graphHandler = {
            param($a)
            $testState.filteredCalls.Add($a)
            & $testState.completeHandler $a
        }
        $filteredOut = Join-Path $TestRoot 'filtered'
        $null = Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $filteredOut
        Assert-True (@($testState.filteredCalls | Where-Object {
            $_.Query -match '^Resources \| where .*summarize'
        }).Count -eq 2) 'Filtered tenant scope runs one initial and one final filtered count, not per-subscription counts.'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -DependencyScope SourceRegionOnly -OutputDirectory $checkPath } 'parameter.*DependencyScope'
        # Execute the actual entry point from a folder containing only the collector.
        $standaloneFolder = Join-Path $TestRoot 'standalone'
        $null = New-Item -ItemType Directory -Path $standaloneFolder
        $standaloneScript = Join-Path $standaloneFolder 'Export-AzureRegionalInventory.ps1'
        Copy-Item -LiteralPath $EntryScript -Destination $standaloneScript
        $testState.graphHandler = $testState.completeHandler
        $standaloneOutput = Join-Path $TestRoot 'standalone-output'
        & {
            # Supply only module availability metadata; all Azure calls remain mocked.
            function Get-Module {
                [CmdletBinding()]param([string[]]$Name, [switch]$ListAvailable)
                foreach ($moduleName in $Name) {
                    if ($moduleName -notin @('Az.Accounts','Az.ResourceGraph','Az.Resources')) {
                        throw "Unexpected module request in standalone test: $moduleName"
                    }
                    [pscustomobject]@{ Name = $moduleName; Version = [version]'1.0.0' }
                }
            }
            $checks = @(& $standaloneScript -SourceRegion southeastasia -SubscriptionId $subA -CheckOnly -OutputDirectory $standaloneOutput)
            Assert-True ($checks.Count -eq 7) 'Standalone script executes preflight without a companion module.'
            Assert-True (-not (Test-Path $standaloneOutput)) 'Standalone CheckOnly does not create inventory files.'
            $standaloneRun = & $standaloneScript -SourceRegion southeastasia -TargetRegion eastasia -SubscriptionId $subA `
                -OutputDirectory $standaloneOutput -RowsPerSheet 1000
            Assert-True ($standaloneRun.IndexedResources -eq 1005 -and $standaloneRun.ImpactedResources -eq 1004) 'Standalone entry point forwards parameters and collects all pages.'
            Assert-True (Test-Path -LiteralPath $standaloneRun.Workbook) 'Standalone entry point writes the workbook.'
            Assert-True ($testState.currentContext -eq $testState.originalContext) 'Standalone execution preserves original Az context.'
        }
        Assert-True (@(Get-ChildItem -LiteralPath $standaloneFolder -File).Count -eq 1) 'Collector distribution consists of one script.'
        Write-Host "PASS: $($testState.passed) assertions. No Azure connection was made."
    } $testRoot ([bool]$Scale) $entryScript
}
finally {
    # Delete only this test's explicitly created, GUID-named temporary directory.
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}
