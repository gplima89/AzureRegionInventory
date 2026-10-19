# Azure regional migration inventory

```powershell
# One-time setup, using your organization's approved package source.
Install-Module Az.Accounts, Az.ResourceGraph, Az.Resources -Scope CurrentUser

# Sign in with Az PowerShell and inventory Canada Central.
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion canadacentral `
    -Login `
    -Verbose
```

Read-only PowerShell 7.2+ collector using Az PowerShell and Azure Resource Graph (ARG).
It generates a real `.xlsx` workbook plus line-delimited JSON companions. It does
not require Azure CLI, Excel, the ImportExcel module, or Python.

## One-file quick start

Download or copy **only [Export-AzureRegionalInventory.ps1](./Export-AzureRegionalInventory.ps1)**.
The collector now includes all query, preflight, normalization, dependency and
Excel-export functions in this one file. No project-specific `.psm1` module,
companion script, repository clone or test runner is needed.

1. Open **64-bit PowerShell 7.2+** and change to the folder where you saved the script.
2. Install the three Az modules through your approved package source (see Setup).
3. Run the command at the top of this README. By default it queries the current
   tenant. Use `-TenantId` to select a tenant, or `-SubscriptionId` to restrict
   collection to one subscription or an array within that tenant.
4. Open the workbook in the new `RegionalInventory-<timestamp>` subfolder under
   your terminal's current directory. Review `Coverage`, `ScopeAudit` and `Errors`
   before treating the result as complete.

Use `-CheckOnly` for a preflight-only run. For parameter descriptions and examples:

```powershell
Get-Help .\Export-AzureRegionalInventory.ps1 -Full
```

The script is organized into collapsible regions for Azure queries, service/
dependency helpers, streaming Excel export, and preflight/orchestration.
Start reading `Invoke-RegionalInventory` for the end-to-end flow; the final line
calls it with the command-line parameters.

## Setup

Install PowerShell 7 and these modules using your approved package source:

```powershell
Install-Module Az.Accounts, Az.ResourceGraph, Az.Resources -Scope CurrentUser
Connect-AzAccount -Tenant '<tenant-guid>'
```

Run in **PowerShell 7**, not Windows PowerShell 5.1. Reader access to each requested
subscription is recommended; narrower RBAC scopes produce a narrower inventory.
The account/subscription list is visibility-based, not an authoritative tenant
subscription census. Supply the client's authoritative subscription list to
detect subscriptions missing from your Az subscription scope. Preflight verifies
the required cmdlet parameters rather than assuming any installed Az version
has compatible commands. Locally validated module versions are Az.Accounts 4.0.2,
Az.ResourceGraph 1.1.0 and Az.Resources 7.8.1; these are tested versions, not
invented minimum requirements.

No commands change Azure resources. `-Login` signs in interactively if requested;
otherwise the script uses your existing `Connect-AzAccount` session. Without
`-SubscriptionId`, ARG requests use `-UseTenantScope` for the current context's
tenant, or the tenant selected by `-TenantId`. With explicit IDs, every request
uses the entire array in `-Subscription`, never one loop per subscription.
All IDs must belong to the selected tenant; run separately for each tenant.
Queries share that tenant's `-DefaultProfile`. Preflight creates one
process-scoped context and restores the original context in a `finally` block;
it does not change your persistent default subscription. For sovereign clouds, sign in
first with the appropriate `Connect-AzAccount -Environment` option and omit
`-Login` on the inventory command. Service principal, managed identity and
other noninteractive sessions can also be established before invocation.

## Automatic prerequisite validation

Every run performs preflight checks before creating inventory files:

- **PowerShell 7.2+** (`#Requires` rejects older versions before script execution),
  64-bit process, and the built-in .NET ZIP/XML types.
- The single collector file includes all project-specific functions; no local
  companion module is needed.
- **Az.Accounts, Az.ResourceGraph and Az.Resources** must be installed and
  importable, with the required authentication/context/query parameters.
  No CLI extension or full Az rollup installation is needed.
- Requested tenant/subscription visibility and Enabled subscription state.
  One `Get-AzSubscription -TenantId` call discovers subscription metadata;
  explicit IDs are validated against this list and deduplicated.
- **One live ARG count query for the whole selected scope**, testing token
  refresh, network access and query access rather than relying on cached
  `Get-AzContext` metadata alone. This is tenant-wide by default, or restricted
  to the explicit subscription array.
- Valid source/destination region names, distinct regions when a target is set,
  a new output path, and a temporary write probe in its nearest existing parent.
  The probe is deleted on close. Final directory/file creation remains subject
  to normal filesystem checks; the probe does not guarantee disk space.

Region-name lookup uses `Get-AzLocation` with the first selected subscription's context.
It validates names only, not destination eligibility for every selected
subscription. ARG inventory/count queries retain the selected tenant or combined
subscription-array scope, regardless of the context's default subscription.

Checks print `[PASS]` messages and stop at the first failure with remediation.
Nothing is automatically installed or upgraded. Successful full runs save the
check results in `Preflight.json`; failed preflight runs do not create an
inventory directory. Authentication and permissions can still change after
preflight, so collection-time failure reporting remains enabled.

To validate prerequisites without collecting an inventory:

```powershell
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion southeastasia `
    -TenantId '<tenant-guid>' `
    -CheckOnly
```

`-CheckOnly` still contacts Azure and briefly writes a uniquely named probe in
the existing output parent, but creates no inventory directory or workbook.
Add `-Login` if you want to sign in first. An existing Azure CLI `az login`
session does **not** authenticate Az PowerShell; use `Connect-AzAccount`.
Live query success proves access only to visible/indexed resources, not complete
tenant RBAC coverage, and does not validate destination deployment readiness.

## Run

Run the self-contained `Export-AzureRegionalInventory.ps1` file. Its command-line
parameters are unchanged from the earlier multi-file version.

```powershell
# Example only: select the actual source and destination for your migration.
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion southeastasia `
    -TargetRegion eastasia `
    -TenantId '<tenant-guid>' `
    -Login `
    -Verbose

# Optional: query an explicit array together instead of the whole tenant.
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion southeastasia `
    -TargetRegion eastasia `
    -TenantId '<tenant-guid>' `
    -SubscriptionId @('<subscription-guid-1>', '<subscription-guid-2>') `
    -OutputDirectory '.\SEA-scoped'

# Optional: faster, narrower discovery, at the cost of inbound-dependency visibility.
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion southeastasia `
    -DependencyScope SourceRegionOnly `
    -OutputDirectory '.\SEA-regional-only'

# See canonical region names (not display names such as "Southeast Asia").
Get-AzLocation | Select-Object Location, DisplayName
```

`TargetRegion` is optional planning context. **It does not validate destination
availability, quotas, physical capacity, policy, residency, or migration support.**
If you omit `-OutputDirectory`, output goes into a new
`.\RegionalInventory-yyyyMMdd-HHmmss` folder under the terminal's **current
working directory** (`Get-Location`), not a fixed `C:\Inventory` path or the
script's directory. For example, running from `D:\Migration` creates
`D:\Migration\RegionalInventory-20261019-113219`.
Use `-OutputDirectory '.\SEA-scoped'` to choose a different relative folder.

The output directory must not already exist; previous inventories are never
overwritten. `-UseDeviceCode` can be combined with `-Login`.

## The 1,000-resource loop

The collector uses **one token chain across the whole selected scope**, not one
per subscription. Resources and resource groups are separate datasets, each with
its own chain:

1. Query `summarize ResourceCount=count()` for the collection scope.
2. Query resources ordered by their unique `id`, retaining `id` in the projection.
3. Request **`Search-AzGraph -First 1000`** with `-UseTenantScope` by default,
   or `-Subscription` containing all explicitly requested IDs.
4. Pass the response's `SkipToken` to **`-SkipToken`** on the next request.
5. Continue until there is no continuation token, even if the initial count has
   already been reached. This avoids stopping early when the inventory changes.
6. Check duplicate IDs, repeated tokens, explicit truncation, the downloaded
   unique count, and a second count query.

For example, 25,342 resources across 20 subscriptions require 26 resource pages,
not 20 separate resource token chains. The same applies if all 25,342 resources
are in one subscription. The two count queries are additional requests.
All queries now execute in-process using `Search-AzGraph`; no `az` subprocess is started.
Tenant scope is limited to the selected tenant and resources visible to the
signed-in identity, not every tenant in the session. The script does not enable
`-AllowPartialScope`: service scope-limit errors must be resolved explicitly.
If needed, pass smaller subscription arrays and document the combined coverage;
dependency resolution cannot cross separate runs.

Continuation tokens are safer than repeatedly issuing `-First 1000` alone,
which would return the first page again. There are no `take`, `limit`, or
`mv-expand` operators in the resource paging query. Dependency extraction happens
locally, avoiding expansion limits and join-related fan-out in ARG.

The core paging follows this pattern, with duplicate/token/count checks around it:

```powershell
$skipToken = $null
do {
    $parameters = @{
        Query = $pageQuery
        First = 1000
        DefaultProfile = $context
        ErrorAction = 'Stop'
    }
    if ($subscriptionId) { $parameters.Subscription = @($subscriptionId) }
    else { $parameters.UseTenantScope = $true }
    if ($skipToken) { $parameters.SkipToken = $skipToken }
    $page = Search-AzGraph @parameters
    # Process $page.Data here, preserving the response object for its metadata.
    $skipToken = $page.SkipToken
} while ($skipToken)
```

Unlike accumulating `$projects += $page`, production processes each page into
the index/JSONL stream, avoiding repeated copying of a growing array. An empty
scope with no matching resources is a valid audited result: the script still
produces a workbook instead of calling `exit`.

Preflight runs one live ARG resource-count request for the entire selected scope.
For `AllAccessibleRegions`, collection reuses that initial count rather than
issuing it again; the final reconciliation count is still queried.
`SourceRegionOnly` has a different filter and therefore needs its own initial
count. Transient throttling, timeout and selected server failures receive bounded
exponential backoff with jitter (six retries by default). Authorization errors
fail that dataset rather than being retried indefinitely. Resource-group and
resource collection are independent: a failed group query is recorded, but does
not prevent resource collection. Any dataset failure makes the run `Partial`.

ARG is eventually consistent and does not offer a transactionally frozen
snapshot for this scan. Equal before/after counts cannot detect replacement of
one resource by another or all concurrent property changes. Collect in a stable
window and rerun discrepancies.

### Performance tips

Az cmdlets reuse the PowerShell process and authentication context. This removes
per-query CLI/Python startup overhead, but module import, authentication, network
latency, throttling and metadata/export size still affect elapsed time.

- For a small test, specify the intended `-SubscriptionId` instead of scanning
  the entire tenant. Metadata discovery still runs once for that tenant;
  ARG collection is restricted to the requested IDs.
- Supply `-TenantId` when known. If omitted, the current Az context's tenant is
  used; the script does not search unrelated signed-in tenants or switch
  automatically to a requested subscription's tenant.
- Tenant and explicit-array runs use one sequential token chain per dataset,
  eliminating repeated count/page loops and context changes per subscription.
- If already signed in, omit `-Login` to avoid another interactive sign-in.
- Use `-Verbose` to see each ARG call's elapsed time. Page progress is always
  printed, including per-page and cumulative resource counts. Stage
  messages distinguish collection, reference resolution and workbook export.
- Modules load once per process, contexts are reused for every page, and
  preflight counts are reused when the query filter matches. Excel column labels
  are cached once per dataset. Completeness checks are not disabled for speed.
- `-DependencyScope SourceRegionOnly` reduces collected metadata, but sacrifices
  inbound dependency visibility from other regions; use only when that trade-off
  is acceptable. The default preserves cross-region discovery.

```powershell
# Reuse your existing Connect-AzAccount session and test one subscription.
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion canadacentral `
    -SubscriptionId '<subscription-guid>' `
    -Verbose
```

## Workbook

| Worksheet | Contents |
|---|---|
| Summary | Snapshot counts by subscription, resource type and impact scope; counts with no tags and no exposed/applicable SKU; formula totals. |
| Inventory | Impacted resource IDs/names/types, groups, subscriptions and tenants, declared/effective regions, normalized SKU, zones, state, identity metadata, tags, selected service details and migration review prompts. |
| Dependencies | Configured outbound and inbound ARM references touching an impacted resource; cross-region/subscription flags, evidence paths and resolution status. |
| DependencyImpact | Per-impacted-resource inbound/outbound reference counts, cross-region references and unresolved references. Counts are evidence rows, not distinct applications or traffic flows. |
| GlobalContext | Global or unspecified-location resources for a separate review; not automatically counted as impacted. |
| ResourceGroups | Group metadata location and tags. Group location is not workload deployment location. |
| ScopeAudit | One row per dataset for the aggregate tenant or subscription-array scope, including tenant/explicit IDs, expected/downloaded counts, page counts and reconciliation status. |
| Coverage | Interpretation, scope, caveats and migration checklist. |
| Errors | Failed dataset/stage, query scope and error message. Empty on a fully successful run. |

`ImpactScope` distinguishes:

- `PrimaryRegion`: ARM `location` matches the source.
- `SecondaryRegion`: a known storage primary/secondary or Cosmos DB location
  includes the source even though the resource's ARM location is elsewhere.
- `InheritedParentRegion`: the resource has no location, but an indexed ARM
  ancestor has the source location. `EffectiveRegion` records this inference
  without changing the original `Region` field.

This is not an exhaustive discovery of replicated data or hidden regional
placement for every Azure service. Global resources and other replication
models require additional service-specific investigation.

### SKU normalization

`SkuName`, `SkuTier`, `SkuCapacity`, `SkuBasis`, and `SkuSourceResourceId`
preserve the meaning and provenance rather than implying common units:

- **VM:** `properties.hardwareProfile.vmSize`, with one VM instance.
- **VM scale set:** top-level SKU and configured instance count.
- **App Service/function app/slot:** inherited App Service plan SKU, tier and
  shared worker count, resolved after all pages/subscriptions are collected.
  Do not sum plan workers once for every app.
- **AKS:** control-plane SKU; individual node-pool VM sizes, counts, min/max,
  zones and subnet references are in `ServiceDetailsJson`. The collector does
  not connect to Kubernetes or inventory pods.
- **Storage/disk:** redundancy or storage-performance SKU, not CPU capacity.
- **SQL:** resource SKU/service objective, pool references, configured size and
  selected HA/backup settings. Pooled DB compute allocation still needs separate
  interpretation; no invented per-database worker count is supplied.
- **PostgreSQL/MySQL, Redis, Cosmos DB, Container Apps, Key Vault:** selected
  service-specific sizing/configuration where exposed by ARG.
- **Other services:** generic ARM SKU when available, with explicit
  `Not exposed / not applicable` otherwise.

Configured disk/database sizes are **not used data measurements**. Missing tags
means no tags at all, not missing specific organizational mandatory tags.
Tags remain JSON so arbitrary tag names do not create thousands of Excel columns.
Resource-group tags are not silently treated as inherited resource tags.

### Dependency interpretation

By default, `AllAccessibleRegions` indexes **all regions of the selected
subscriptions**, not just the source. This is intentional: otherwise an app or
network resource in another region pointing into the source would be missed.
Only relationships touching impacted resources are exported.

The collector scans configuration and identity objects for exact ARM resource-ID
values and user-assigned identity ID keys. It also adds ARM parent relationships
and AKS node-resource-group membership as explicitly labeled structural
inferences. It resolves references after all pages in the selected scope have been scanned.
Subnets and other unindexed child IDs can resolve to their nearest indexed
ancestor; the exact referenced ID is preserved alongside that ancestor.

Examples include VM/NIC/disk/VNet references, app/plan links, managed identities,
private endpoints, peering and configured linked service IDs where ARG exposes
them. A configuration reference does not prove runtime traffic, dependency
criticality, or a valid migration sequence.

This is a **one-hop metadata analysis**, not a transitive application dependency
graph. Unresolved IDs may be outside scope, invisible under RBAC, unsupported by
ARG, deleted, or a reference to an unindexed subresource. They are not automatically
broken references. Cross-region status remains unspecified for resources with
no declared region, even where a regional parent was inferred in Inventory.

`SourceRegionOnly` filters before indexing, so inbound references from unrelated
regions, late shared-plan resolution and some target resolution can be missed.
It includes global/blank locations and known storage/Cosmos region fields as
additional context, but does not guarantee tenant-wide dependency coverage.

## Scale and output handling

- 64-bit PowerShell is recommended. Start with a representative subscription,
  then expand. Hundreds of thousands of resources can take a long time, depending
  on metadata size, permissions, throttling and the number of subscriptions.
- Collection is sequential to avoid unbounded tenant API load.
- Full resource properties exist only in the current ARG page in memory.
  The persisted staging/output uses selected normalized metadata, not complete
  arbitrary property bags, app settings, passwords, keys or secret values.
- A compact resource-ID/SKU/location index stays in RAM for cross-region
  resolution. Memory is **not constant**; budget several GB for very large
  estates and measure the representative run. Excel cells themselves are streamed
  directly into the compressed XLSX rather than held in an in-memory spreadsheet.
- Worksheets split at 250,000 data rows by default, and never exceed Excel's
  1,048,576 total-row limit. Use `-RowsPerSheet` to adjust the split.
  Summary must fit on one sheet so its total formulas remain valid; the script
  fails explicitly if this setting is too small for the number of summary groups.
- Cells exceeding Excel's 32,767-character limit are marked truncated. JSONL
  companions preserve the complete selected values. Strings are written as
  literal text, not interpreted as Excel formulas.
- Summary groups are measured snapshot aggregates, not live formulas linked to
  individual Inventory rows. Total formulas update when summary group counts
  change, but editing inventory rows does not rebuild the summary groups.
- JSONL companions are suitable for subsequent Power BI, scripting or database
  ingestion. Do not discard them merely because the workbook opens successfully.

Treat **all outputs as confidential**, including tags, names, addresses and
identity IDs. Tags and other user-controlled metadata can themselves contain
sensitive information even though secret APIs are never called. Use an
approved access-controlled output directory, not a broadly shared sync folder.

`Run.json` records `Collecting`, `Completed`, `Partial`, or `Failed`. On a
subscription failure, collection continues, a clearly marked partial workbook
is saved, and the script throws a terminating error. Partial rows from failed
subscriptions can be present; do not use them as a complete inventory.
Subscriptions failing preflight stop the entire run before inventory creation;
the partial-workbook behavior applies to failures after preflight succeeds.
If export fails or the process is interrupted, an `.xlsx.partial` file and JSONL
staging can remain for troubleshooting. No final workbook name is published
until export succeeds. There is no resume capability; use a new output path
for a fresh run.

## What to add for migration decisions

Inventory is a starting point, not a migration assessment certification.
Supplement it with:

1. **Owner/application mapping:** agree tags, business criticality, maintenance
   windows, RTO/RPO, data residency and migration waves with workload owners.
2. **Runtime dependencies:** Azure Migrate dependency analysis where supported,
   Application Insights application maps, approved network-flow telemetry and
   workload-owner validation for DNS, FQDN, connection-string, SaaS and on-premises
   consumers not visible as ARM IDs.
3. **Target feasibility:** VM SKU/zone availability, subscription quotas,
   capacity reservations, PaaS regional support, provider registration, policies
   and region-specific feature differences. VM SKU listings are not a guarantee
   of physical deployment capacity.
4. **Data/state:** actual data volumes and change rates, database replication,
   storage transfer strategy, disk replication, backup/restore tests, encryption
   and CMK access, failover groups and service-specific migration tooling.
5. **Cutover:** public/private DNS, IP allowlists, private endpoints, ExpressRoute/
   VPN, routes/firewalls, certificates, resource locks, policies, diagnostics,
   managed identities/RBAC, backup coverage and rollback.
6. **Commercial impact:** utilization, reservations/savings plans, target SKU
   sizing, egress/transfer and temporary parallel-running cost.

These additional sources are intentionally not represented as successfully
collected facts. The script does not call secret, data-plane, cost, monitoring,
policy, lock or role-assignment APIs.

## Troubleshooting

### Start with preflight

Run from PowerShell 7 in the folder containing the collector script:

```powershell
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion southeastasia `
    -TenantId '<tenant-guid>' `
    -CheckOnly `
    -Verbose
```

Preflight stops at the first failure. Correct it and rerun until all checks pass.
If your session needs authentication, add `-Login`. Use a new output folder for
each subsequent collection attempt; the collector cannot resume an interrupted run.

### Setup and authentication errors

| Error or symptom | Likely cause | Fix |
|---|---|---|
| Script requires PowerShell 7.2, or a parameter such as `AsHashtable` is not recognized | Running Windows PowerShell 5.1 or an older runtime | Open `pwsh` and check `$PSVersionTable.PSVersion`. Install or update PowerShell 7 through your approved process if needed. |
| `64-bit PowerShell is required` | A 32-bit terminal/process is running | Start 64-bit PowerShell 7; `[Environment]::Is64BitProcess` should return `True`. |
| `Required module file missing` referring to `RegionalInventory.psm1` | An older multi-file entry script is running | Replace it with the current self-contained `Export-AzureRegionalInventory.ps1`. No project-specific module file is needed; the three Az modules are still required. |
| Script execution is blocked or the file is not digitally signed | Execution policy, downloaded-file marking, or organizational application control | Inspect `Get-ExecutionPolicy -List`. Use your approved signing/unblocking process or contact your administrator. Do not bypass organizational policy or disable security controls. |
| `Required module Az... is missing`, or a cmdlet is not recognized | Required Az module is missing or cannot be imported | Install `Az.Accounts`, `Az.ResourceGraph` and `Az.Resources` using your approved source, for example `Install-Module Az.Accounts, Az.ResourceGraph, Az.Resources -Scope CurrentUser`. No CLI extension is used. |
| A required cmdlet parameter such as `SkipToken` is missing | An incompatible Az module version is loaded | Check `Get-Module Az.*`; update the owning module through your approved process and open a fresh PowerShell terminal to avoid mixing loaded assembly versions. |
| `No Az PowerShell session`, subscription discovery failed, or an AADSTS/token-expiration error | Missing or expired Az authentication, wrong tenant, or Conditional Access requirements | Run `Connect-AzAccount -Tenant '<tenant-guid>'`, complete authentication, then rerun preflight. Azure CLI `az login` is a separate session and is insufficient. |
| Browser-based login cannot complete | Browser access or authentication flow is restricted | If your organization permits device-code authentication, use `-Login -UseDeviceCode`. If Conditional Access blocks it, use an approved authentication flow with your administrator. |
| `UseDeviceCode requires Login` | Device-code option was supplied without requesting sign-in | Supply both switches: `-Login -UseDeviceCode`. |
| `Subscription not visible`, discovery failed or `No subscriptions in the requested scope` | Incorrect tenant/ID, missing subscription visibility, or stale authentication | Sign in to the correct tenant, inspect `Get-AzSubscription`, and verify requested GUIDs and Reader access. Do not remove required subscriptions just to obtain a passing run. |
| `Cross-tenant arrays are not supported` | Requested IDs are not all visible in the current/selected tenant | Select the correct `-TenantId` and verify access. For IDs belonging to different tenants, run separately per tenant and record the dependency-analysis boundary. The script does not widen scope automatically. |
| Subscription is `Disabled` or otherwise not `Enabled` | Subscription state does not permit this collection | Ask the subscription administrator to investigate. Exclude it only if the inventory owner explicitly accepts the reduced scope and records the gap. |

### Azure query and pagination errors

| Error or symptom | Likely cause | Fix |
|---|---|---|
| `unrecognized arguments: --subscription ...` from an old CLI-based collector | An older copy of the script is running | Replace it with the current self-contained script, which calls `Get-AzLocation` and `Search-AzGraph` rather than Azure CLI, then rerun the script. |
| `Live ARG preflight failed`, `AuthorizationFailed` or HTTP 403 | Token/tenant mismatch, insufficient RBAC, or another access restriction | Read the original error included in the message. Refresh login for the subscription's tenant and have the administrator verify Reader access at the intended scope. After a role change, allow propagation and rerun preflight. |
| Tenant subscription-scope limit error | The requested tenant scope exceeds ARG's supported subscription scope | Use explicit smaller `-SubscriptionId` arrays, collecting each to a new output directory, and reconcile them against the intended estate. Partial scopes are deliberately disabled; do not enable `AllowPartialScope` to conceal missing subscriptions. Separate runs cannot resolve dependencies across their boundaries. |
| SSL/certificate, proxy, DNS, connection or timeout errors | Corporate proxy/VPN/firewall or certificate trust is preventing Azure access | Confirm access to the configured Azure cloud through your approved network. Ask IT to verify PowerShell/.NET proxy configuration and trusted enterprise certificates. Do not disable TLS verification. |
| `429`, `TooManyRequests`, or transient 502/503/504 warnings | Azure throttling or a transient service failure | Allow the built-in bounded retries to finish. If exhausted, reduce competing inventory jobs, wait and rerun with a new output path. Splitting subscription batches can reduce competing load; keep records of the full intended scope. |
| `Unknown source region` / `Unknown target region` | A display name or misspelled region was supplied | Use the canonical `Location` from `Get-AzLocation`, such as `southeastasia`, not `Southeast Asia`. |
| `SourceRegion and TargetRegion must be different` | The same region was supplied twice | Correct the destination, or omit `-TargetRegion` when only gathering inventory. |
| `ARG count mismatch` or `Duplicate resource ID while paging` | Resources changed during collection, ARG indexing changed, or paging was inconsistent | Rerun during a quieter window using a new output folder. Keep the error/count evidence if it recurs. Do not disable reconciliation to force a successful result. |
| Repeated continuation token, empty page with a token, or count mismatch after a missing token | Unexpected ARG response or a modified query that is not pageable | Check Az.ResourceGraph compatibility. Retain scalar `id`, stable ordering and token handling; remove `take`, `limit` or `sample` operators. Never treat the first 1,000 rows as the complete result. |
| Invalid ARG count response or `Data and SkipToken are required` | Incompatible tooling or a custom query/response-shaping change | Check Az module versions and reproduce the small count query below. Preserve the response object until `SkipToken` is read; do not enumerate or project it away. Never replace failed responses with empty inventories. |

To isolate a live query problem without running the entire collector:

```powershell
$originalContext = Get-AzContext
try {
    $context = Set-AzContext -Subscription '<subscription-guid>' -Tenant '<tenant-guid>' -Scope Process
    Search-AzGraph `
        -Subscription '<subscription-guid>' `
        -DefaultProfile $context `
        -First 1 `
        -Query 'Resources | summarize ResourceCount=count()'
}
finally {
    Set-AzContext -Context $originalContext -Scope Process | Out-Null
}
```

A successful response does not prove visibility of every resource in the
subscription; compare the scope against the client's authoritative inventory.

### Output, scale and Excel errors

| Error or symptom | Likely cause | Fix |
|---|---|---|
| `Output directory already exists` | A previous run used that path | Omit `-OutputDirectory` to use a new timestamped folder under the current directory, or choose a new relative path such as `.\SEA-retry-01`. Keep previous evidence rather than overwriting it. |
| `Output parent is not writable`, access denied, or path errors | Filesystem permissions, an invalid path or a protected folder | Choose an approved, access-controlled local filesystem path. Check `Get-Location` to confirm where relative paths resolve. A successful parent probe does not guarantee permissions for every later file operation. |
| Export fails with a file-lock, I/O or disk-space error | Open handles, sync software, insufficient space or storage failure | Close applications holding the affected output, check free space and use an approved nonsynchronized folder if sync interference is suspected. Retry to a new path; do not delete unrelated files or previous inventories. |
| Memory exhaustion or very slow collection | Large resource index, large metadata, limited RAM or API throttling | Use 64-bit PowerShell with sufficient RAM and close unnecessary local workloads. Start with representative subscriptions. Batching subscriptions reduces memory but loses dependency resolution across batches; `SourceRegionOnly` also reduces inbound visibility. Document these trade-offs. Lowering worksheet split size does not eliminate the in-memory resource index. |
| `Summary needs ... data rows; increase RowsPerSheet` | Worksheet split setting is smaller than the summary group count plus its total row | Increase `-RowsPerSheet`, normally back to the default `250000`, without exceeding `1048575`. Rerun to a new output directory. |
| Only an `.xlsx.partial` file exists, or Excel cannot open it | Export was interrupted or failed | Check `Run.json` and the terminal error. The partial ZIP is not a completed workbook; do not rename it to imply success. Preserve JSONL files for investigation and rerun. |
| Excel displays a repair warning for a final workbook | File corruption, incomplete transfer, or an export defect | Preserve the original file and any repair report. Compare with the original local output, rerun the offline tests and retry the collection to a new folder. Do not assume repaired or missing rows are complete inventory data. |
| An independent reader reports `Package is an OLE compound document` or asks for a password | The stored workbook may be encrypted or information-protected; confirm with your organization's tooling | Open it in Excel using your authorized account and check its protection/sensitivity status. Libraries without protection support cannot validate an encrypted package. Do not strip protection, rename containers or bypass policy to make a test pass. |
| A cell ends with `[TRUNCATED: see JSONL companion]` | The selected metadata exceeded Excel's cell text limit | Read the corresponding resource/row in the JSONL companion. This is an explicit display limit, not evidence that the entire collection stopped. |

### Incomplete or unexpected results

- **`PARTIAL inventory saved`:** inspect the `Errors` and `ScopeAudit` worksheets,
  then fix the reported dataset/scope failures. Some rows may already have been
  saved from a failed token chain. Do not present that workbook as a complete
  estate inventory; rerun to a new output folder after remediation.
- **`Run.json` says `Collecting` after the process has stopped:** the process
  likely terminated before it could record completion. `Failed` records a
  handled failure; `Partial` records collection failures with an exported
  workbook. Investigate before using any files.
- **Zero or fewer resources than expected:** check source region, tenant,
  explicit subscription list, RBAC visibility, ARG indexing delay and supported
  resource types. Resource-group metadata location does not determine the
  location of its resources. Count reconciliation validates the queried visible
  scope, not the client's full estate.
- **App SKU unknown:** confirm the referenced App Service plan is visible and
  included in the subscription scope. Prefer `AllAccessibleRegions` for
  cross-region shared plans. Some services genuinely have no exposed resource SKU.
- **Unresolved dependencies or no dependency rows:** confirm both ends are
  visible and use `AllAccessibleRegions` for inbound discovery. Not all
  subresources are independently indexed. FQDNs, connection strings and runtime
  traffic are intentionally not discovered; no rows does not mean no dependencies.
- **Resources appear in `GlobalContext` instead of `Inventory`:** global or
  regionless resources are not automatically source-region workloads. Review
  global routing and unspecified placement manually; do not count every global
  resource as impacted without evidence.

### Collecting safe support information

Record the failing stage, sanitized error message, UTC time, Az module/runtime
versions, and whether the issue occurs in `-CheckOnly`. For collection failures,
also inspect `Run.json`, `Preflight.json` when present, and the relevant
`ScopeAudit`/`Errors` rows.

```powershell
$PSVersionTable.PSVersion
[Environment]::Is64BitProcess
Get-Location
Get-Module -ListAvailable Az.Accounts, Az.ResourceGraph, Az.Resources |
    Select-Object Name, Version
Get-Command Search-AzGraph, Get-AzSubscription, Get-AzLocation |
    Select-Object Name, Source
```

Sanitize subscription/tenant IDs, resource names, tags, paths and other client
metadata before sharing outside the approved support channel. Never share
access tokens, credentials, secret values or unreviewed debug logs. The
diagnostic commands above do not request token values.

## Offline verification

This section is for maintainers. End users do not need the test script to run
the inventory. To run the development suite, keep
`Test-RegionalInventory.ps1` beside the single collector:

```powershell
.\Test-RegionalInventory.ps1 -Scale
```

Tests use mocked Az cmdlets and make **no Azure connection**. They cover
preflight-only mode, module/session checks, live-query failures,
disabled or missing subscriptions, invalid regions and output paths, plus a
100,001-resource scope, >1,000-resource multi-subscription orchestration, continuation tokens,
empty/exact-page scopes, duplicates, malformed responses, count mismatches, retry behavior,
tenant-default and combined-array scope/context forwarding, cross-tenant array
rejection, context restoration,
service-specific SKUs, late cross-subscription plan resolution, secondary-region
storage, inherited child location, inbound/cross-region references, partial
failures, worksheet splitting, XML validity, formula injection prevention,
text limits, and cached formula totals.
The tests also copy only the collector into an isolated folder and execute its
real preflight and full entry point with mocked Azure calls. This verifies that
neither a companion module nor the test script is needed to run the collector.
Helper functions are loaded from the collector's parsed syntax tree for isolated
unit tests, without triggering a live inventory.
If ImportExcel is already installed, tests additionally open the workbook with
its independent EPPlus reader and recalculate summary totals. ImportExcel is not
a production dependency. No dependencies are installed by the tests.

## References

- [ARG large datasets, page size and continuation tokens](https://learn.microsoft.com/en-us/azure/governance/resource-graph/concepts/work-with-data)
- [Search-AzGraph parameters and paging](https://learn.microsoft.com/en-us/powershell/module/az.resourcegraph/search-azgraph)
- [Connect-AzAccount authentication](https://learn.microsoft.com/en-us/powershell/module/az.accounts/connect-azaccount)
- [ARG supported resource types and tables](https://learn.microsoft.com/en-us/azure/governance/resource-graph/reference/supported-tables-resources)
