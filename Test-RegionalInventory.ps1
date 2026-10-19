#Requires -Version 7.2
[CmdletBinding()]
param([switch]$Scale)
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'RegionalInventory.psm1') -Force
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ("RegionalInventoryTests-" + [guid]::NewGuid())
$null = New-Item -ItemType Directory -Path $testRoot
try {
    & (Get-Module RegionalInventory) {
        param($TestRoot, $RunScale)
        $script:passed = 0
        function Assert-True {
            param([bool]$Condition, [string]$Message)
            if (-not $Condition) { throw "ASSERTION FAILED: $Message" }
            $script:passed++
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
        $script:realAz = ${function:Invoke-InventoryAz}
        function Invoke-InventoryAz { param([string[]]$Arguments) & $script:azHandler $Arguments }
        $subA = '11111111-1111-1111-1111-111111111111'
        $subB = '22222222-2222-2222-2222-222222222222'
        $prefixA = "/subscriptions/$subA/resourceGroups/workload/providers"
        $prefixB = "/subscriptions/$subB/resourceGroups/shared/providers"
        $script:pageCount = 2505
        $script:pageCalls = [Collections.Generic.List[object]]::new()
        $script:azHandler = {
            param($a)
            $q = $a[[array]::IndexOf($a, '-q') + 1]
            if ($q -match 'summarize') { return @{ data = @(@{ ResourceCount = $script:pageCount }) } }
            $script:pageCalls.Add($a)
            $first = [int]$a[[array]::IndexOf($a, '--first') + 1]
            $start = if ('--skip-token' -in $a) { [int]$a[[array]::IndexOf($a, '--skip-token') + 1] } else { 0 }
            $end = [Math]::Min($start + $first, $script:pageCount)
            $rows = @(for ($i = $start; $i -lt $end; $i++) { @{ id = "/resource/$i" } })
            @{ data = $rows; skip_token = $(if ($end -lt $script:pageCount) { [string]$end } else { $null }) }
        }
        $script:received = 0L
        $result = Invoke-ArgPaged $subA 'Resources' 'id' { param($rows) $script:received += $rows.Count }
        Assert-True ($result.Downloaded -eq 2505 -and $result.Pages -eq 3) '2505 resources require three pages.'
        Assert-True ($script:received -eq 2505) 'Every row reaches callback.'
        Assert-True ($script:pageCalls[1] -contains '--skip-token') 'Second page uses continuation token.'
        Assert-True (($script:pageCalls[0] -join ' ') -match '--first 1000') 'First is exactly 1000.'
        $script:pageCount = 0
        $result = Invoke-ArgPaged $subA 'Resources' 'id' { throw 'No rows expected' }
        Assert-True ($result.Downloaded -eq 0) 'Empty scope succeeds.'
        $script:pageCount = 2000
        $result = Invoke-ArgPaged $subA 'Resources' 'id' { param($rows) }
        Assert-True ($result.Pages -eq 2) 'Exact multiple of 1000 completes without an extra page.'
        if ($RunScale) {
            $script:pageCalls.Clear()
            $script:pageCount = 100001
            $result = Invoke-ArgPaged $subA 'Resources' 'id' { param($rows) }
            Assert-True ($result.Downloaded -eq 100001 -and $result.Pages -eq 101) '100001 resources in one subscription.'
        }
        $script:azHandler = {
            param($a)
            if (($a -join ' ') -match 'summarize') { return @{ data = @(@{ ResourceCount = 2 }) } }
            @{ data = @(@{ id = 'one' }); resultTruncated = 'true' }
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'truncated'
        $script:azHandler = {
            param($a)
            if (($a -join ' ') -match 'summarize') { return @{ data = @(@{ ResourceCount = 2 }) } }
            @{ data = @(@{ id = 'one' }); skip_token = 'same' }
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'Duplicate'
        $script:azHandler = {
            param($a)
            if (($a -join ' ') -match 'summarize') { return @{ data = @(@{ ResourceCount = 2 }) } }
            @{ data = @(@{ id = 'one' }); skip_token = $null }
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'count mismatch'
        $script:mockPage = 0
        $script:azHandler = {
            param($a)
            if (($a -join ' ') -match 'summarize') { return @{ data = @(@{ ResourceCount = 3 }) } }
            $script:mockPage++
            @{ data = @(@{ id = "unique-$script:mockPage" }); skip_token = 'repeated' }
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'repeated continuation'
        $script:mockCounts = 0
        $script:azHandler = {
            param($a)
            if (($a -join ' ') -match 'summarize') {
                $script:mockCounts++
                return @{ data = @(@{ ResourceCount = $script:mockCounts }) }
            }
            @{ data = @(@{ id = 'one' }) }
        }
        Assert-Throws { Invoke-ArgPaged $subA 'Resources' 'id' {} } 'before=1, downloaded=1, after=2'

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
        $refs = @(Get-ArmReferences -Value @{
            subnet = @{ id = "$prefixA/Microsoft.Network/virtualNetworks/vnet/subnets/default" }
            credential = "Password=secret;Resource=$planId"
            identities = @{ $planId = @{} }
        })
        Assert-True ($refs.Count -eq 2) 'Only exact ARM values or identity keys; no connection string matching.'
        Assert-True ((Resolve-IndexedResource "$planId/child/one" $index).Id -eq $planId) 'Subresource ancestor resolves.'

        # Exercise the real CLI retry wrapper without invoking a native Azure command.
        $script:nativeAttempts = 0
        function az {
            $script:nativeAttempts++
            if ($script:nativeAttempts -eq 1) {
                $global:LASTEXITCODE = 1
                # A PowerShell error stream is redirected by the wrapper just like native stderr.
                Write-Error 'TooManyRequests 429' -ErrorAction Continue
            }
            else { $global:LASTEXITCODE = 0; '{"ok":true}' }
        }
        function Start-Sleep { param($Seconds) }
        $retry = & $script:realAz -Arguments @('mock') -MaxRetries 1
        Assert-True ($retry.ok -and $script:nativeAttempts -eq 2) 'Throttle is retried.'
        function az { $global:LASTEXITCODE = 1; Write-Error 'AuthorizationFailed' -ErrorAction Continue }
        Assert-Throws { & $script:realAz -Arguments @('mock') } 'AuthorizationFailed'

        # Full offline orchestration: >1000 resources, late cross-subscription plan, inbound reference and secondary region.
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
            New-TestResource $planId 'microsoft.web/serverfarms' 'eastasia' @{} @{ name = 'P1v3'; tier = 'PremiumV3'; capacity = 3 }
            New-TestResource "$prefixB/Microsoft.Example/consumers/consumer" 'microsoft.example/consumers' 'eastasia' @{ upstreamId = $appId }
        )
        $script:testResources = @{ $subA = @($resourcesA.ToArray()); $subB = $resourcesB }
        $script:simulateFailure = $false
        $script:azHandler = {
            param($a)
            if ($a[0] -eq 'version') { return @{ 'azure-cli' = '2.80.0' } }
            switch ("$($a[0]) $($a[1])") {
                'account show' { return @{ id = $subA } }
                'extension list' { return @(@{ name = 'resource-graph' }) }
                'account list' { return @(
                    @{ id = $subA; name = 'Workloads'; tenantId = 'tenant'; state = 'Enabled' }
                    @{ id = $subB; name = 'Shared'; tenantId = 'tenant'; state = 'Enabled' }
                ) }
                'account list-locations' { return @(@{ name = 'southeastasia' }, @{ name = 'eastasia' }) }
                'graph query' {
                    $sub = $a[[array]::IndexOf($a, '--subscriptions') + 1]
                    $q = $a[[array]::IndexOf($a, '-q') + 1]
                    if ($script:simulateFailure -and $sub -eq $subB -and $q -match '^ResourceContainers') { throw 'AuthorizationFailed test fixture' }
                    $all = @(if ($q -match '^ResourceContainers') {
                        @{ id = "/subscriptions/$sub/resourceGroups/metadata-group"; name = 'metadata-group'
                            subscriptionId = $sub; location = 'eastasia'; tags = @{ Owner = 'Test' } }
                    } else { $script:testResources[$sub] })
                    if ($q -match 'summarize') { return @{ data = @(@{ ResourceCount = $all.Count }) } }
                    $start = if ('--skip-token' -in $a) { [int]$a[[array]::IndexOf($a, '--skip-token') + 1] } else { 0 }
                    $last = [Math]::Min($start + 1000, $all.Count)
                    $data = @()
                    if ($all.Count) { $data = @($all[$start..($last - 1)]) }
                    return @{ data = $data; skip_token = $(if ($last -lt $all.Count) { [string]$last } else { $null }) }
                }
                default { throw "Unexpected mock CLI call: $($a -join ' ')" }
            }
        }
        $script:completeHandler = $script:azHandler
        $checkPath = Join-Path $TestRoot 'preflight-only'
        $checks = @(Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $checkPath -CheckOnly)
        Assert-True ($checks.Count -eq 9) 'CheckOnly returns runtime, CLI, extension, scope, per-subscription live access, regions and path checks.'
        Assert-True (-not (Test-Path $checkPath)) 'CheckOnly does not create the output directory.'
        Assert-True (@(Get-ChildItem -LiteralPath $TestRoot -Filter '.inventory-preflight-*').Count -eq 0) 'Write probe is cleaned up.'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $checkPath -CheckOnly -UseDeviceCode } 'requires Login'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $TestRoot -CheckOnly } 'already exists'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion invalidregion -OutputDirectory $checkPath -CheckOnly } 'Unknown source region'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -TargetRegion southeastasia -OutputDirectory $checkPath -CheckOnly } 'must be different'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -SubscriptionId 'not-visible' -OutputDirectory $checkPath -CheckOnly } 'not in CLI account list'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -TenantId 'another-tenant' -OutputDirectory $checkPath -CheckOnly } 'No subscriptions'
        $script:preflightScenario = ''
        $script:azHandler = {
            param($a)
            if ($a[0] -eq 'version' -and $script:preflightScenario -eq 'OldCli') { return @{ 'azure-cli' = '2.21.0' } }
            if ($a[0] -eq 'version' -and $script:preflightScenario -eq 'InvalidCli') { return @{ 'azure-cli' = 'invalid' } }
            if ($a[0] -eq 'extension' -and $script:preflightScenario -eq 'MissingExtension') { return @() }
            if ($a[0] -eq 'account' -and $a[1] -eq 'show' -and $script:preflightScenario -eq 'LoggedOut') { throw 'Please run az login' }
            if ($a[0] -eq 'graph' -and $script:preflightScenario -eq 'ExpiredToken') { throw 'AADSTS token expired' }
            if ($a[0] -eq 'account' -and $a[1] -eq 'list' -and $script:preflightScenario -eq 'Disabled') {
                return @{ id = $subA; name = 'Disabled'; tenantId = 'tenant'; state = 'Disabled' }
            }
            & $script:completeHandler $a
        }
        foreach ($case in @(
            @('OldCli', '2.22.0'), @('InvalidCli', '2.22.0'),
            @('MissingExtension', 'az extension add'), @('LoggedOut', 'az login'),
            @('ExpiredToken', 'Live ARG preflight failed'), @('Disabled', 'not Enabled')
        )) {
            $script:preflightScenario = $case[0]
            Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $checkPath -CheckOnly } $case[1]
            Assert-True (-not (Test-Path $checkPath)) "Preflight failure $($case[0]) stops before output collection."
        }
        $script:azHandler = $script:completeHandler
        $out = Join-Path $TestRoot 'complete'
        $run = Invoke-RegionalInventory -SourceRegion southeastasia -TargetRegion eastasia -OutputDirectory $out -RowsPerSheet 1000
        Assert-True ($run.ImpactedResources -eq 1004 -and $run.IndexedResources -eq 1007) 'Full run selects primary, secondary and inherited resources, not every region.'
        $inventory = @(Get-JsonLines (Join-Path $out 'Inventory.jsonl'))
        $appRow = @($inventory | Where-Object { $_.ResourceId -eq $appId })[0]
        Assert-True ($appRow.SkuName -eq 'P1v3') 'Late plan resolution updates final app inventory.'
        Assert-True (@(Get-JsonLines (Join-Path $out 'ResourceGroups.jsonl')).Count -eq 2) 'Resource-group dataset collected independently of workload region.'
        $slotRow = @($inventory | Where-Object { $_.ResourceType -eq 'microsoft.web/sites/slots' })[0]
        Assert-True ($slotRow.ImpactScope -eq 'InheritedParentRegion' -and $slotRow.EffectiveRegion -eq 'southeastasia') 'Regionless child inherits indexed parent region.'
        $edges = @(Get-JsonLines (Join-Path $out 'Dependencies.jsonl'))
        Assert-True (@($edges | Where-Object { $_.ImpactDirection -eq 'Inbound to impacted' }).Count -eq 1) 'Inbound consumers from other regions included.'
        Assert-True (@($edges | Where-Object { $_.CrossRegion -eq 'Yes' }).Count -eq 2) 'Known cross-region references are flagged; blank child locations remain unspecified.'
        Assert-True (-not (Test-Path (Join-Path $out 'InventoryStaging.jsonl'))) 'Staging cleaned after success.'
        $manifest = Get-Content (Join-Path $out 'Run.json') -Raw | ConvertFrom-Json
        Assert-True ($manifest.Status -eq 'Completed') 'Successful manifest.'
        Assert-True (Test-Path (Join-Path $out 'Preflight.json')) 'Successful preflight checks saved alongside inventory.'

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
            $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
            $reader = [IO.StreamReader]::new($entry.Open())
            try { $summary = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
            $ns = [Xml.XmlNamespaceManager]::new($summary.NameTable)
            $ns.AddNamespace('s', $summary.DocumentElement.NamespaceURI)
            $totals = $summary.SelectNodes('//s:c[s:f]', $ns)
            Assert-True ($totals.Count -eq 3 -and $totals[0].v -eq '1004') 'Summary formula totals have correct cached values.'
        }
        finally { $zip.Dispose() }
        if (Get-Module -ListAvailable ImportExcel) {
            Import-Module ImportExcel
            $package = Open-ExcelPackage -Path $run.Workbook
            try {
                Assert-True ($package.Workbook.Worksheets['Inventory'].Dimension.Rows -eq 1001) 'Independent EPPlus reader opens first inventory worksheet.'
                Assert-True ($package.Workbook.Worksheets['Inventory_2'].Dimension.Rows -eq 5) 'Independent reader sees exact split remainder.'
                Assert-True ($package.Workbook.Worksheets['Summary'].Cells['E6'].Value -eq 1004) 'Independent reader sees cached total.'
                [OfficeOpenXml.CalculationExtension]::Calculate($package.Workbook)
                Assert-True ($package.Workbook.Worksheets['Summary'].Cells['E6'].Value -eq 1004) 'Independent engine recalculates formula total.'
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
        $script:simulateFailure = $true
        $partialOut = Join-Path $TestRoot 'partial'
        Assert-Throws { Invoke-RegionalInventory -SourceRegion southeastasia -OutputDirectory $partialOut } 'PARTIAL inventory saved'
        $partial = Get-Content (Join-Path $partialOut 'Run.json') -Raw | ConvertFrom-Json
        Assert-True ($partial.Status -eq 'Partial' -and (Test-Path (Join-Path $partialOut 'RegionalInventory.xlsx'))) 'Partial workbook preserved with explicit failure status.'
        Write-Host "PASS: $script:passed assertions. No Azure connection was made."
    } $testRoot ([bool]$Scale)
}
finally {
    # Delete only this test's explicitly created, GUID-named temporary directory.
    Remove-Item -LiteralPath $testRoot -Recurse -Force
}
