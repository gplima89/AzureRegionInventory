# Azure regional migration inventory

Read-only PowerShell 7.2+ collector using Azure CLI and Azure Resource Graph (ARG).
It generates a real `.xlsx` workbook plus line-delimited JSON companions. It does
not require Excel, the ImportExcel module, Az PowerShell modules, or Python.

## Setup

Install PowerShell 7 and Azure CLI using your approved installation method, then:

```powershell
az extension add --name resource-graph
az login --tenant '<tenant-guid>'
```

Run in **PowerShell 7**, not Windows PowerShell 5.1. Reader access to each requested
subscription is recommended; narrower RBAC scopes produce a narrower inventory.
The account/subscription list is visibility-based, not an authoritative tenant
subscription census. Supply the client's authoritative subscription list to
detect subscriptions missing from your CLI account list.

No commands change Azure resources. `-Login` signs in interactively if requested;
otherwise the script uses your existing CLI authentication. `--subscription` is
passed on each ARG request so its authentication context matches that
subscription, including when the account list spans tenants. The script does
not change the persistent default subscription. For sovereign clouds, configure
the Azure CLI cloud and sign in there before running.

## Automatic prerequisite validation

Every run performs preflight checks before creating inventory files:

- **PowerShell 7.2+** (`#Requires` rejects older versions before script execution),
  64-bit process, and the built-in .NET ZIP/XML types.
- The companion module must be beside the entry script. No Az, ImportExcel,
  Excel, Python, or PowerShell Gallery modules are required for production.
- **Azure CLI 2.22.0+** on PATH and the installed `resource-graph` extension.
- CLI session/account refresh, requested tenant/subscription visibility and
  Enabled subscription state.
- A **live ARG count query for every selected subscription**, testing token
  refresh, network access and query access rather than relying on cached
  `az account show` metadata alone.
- Valid source/destination region names, distinct regions when a target is set,
  a new output path, and a temporary write probe in its nearest existing parent.
  The probe is deleted on close. Final directory/file creation remains subject
  to normal filesystem checks; the probe does not guarantee disk space.

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
Add `-Login` if you want to sign in first. An existing Az PowerShell
`Connect-AzAccount` session alone does **not** authenticate Azure CLI; use `az login`.
Live query success proves access only to visible/indexed resources, not complete
tenant RBAC coverage, and does not validate destination deployment readiness.

## Run

Keep the entry script and module in the same directory.

```powershell
# Example only: select the actual source and destination for your migration.
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion southeastasia `
    -TargetRegion eastasia `
    -TenantId '<tenant-guid>' `
    -Login `
    -Verbose

# Prefer an explicit, authoritative subscription scope for a large client.
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion southeastasia `
    -TargetRegion eastasia `
    -SubscriptionId @('<subscription-guid-1>', '<subscription-guid-2>') `
    -OutputDirectory '.\SEA-scoped'

# Optional: faster, narrower discovery, at the cost of inbound-dependency visibility.
.\Export-AzureRegionalInventory.ps1 `
    -SourceRegion southeastasia `
    -DependencyScope SourceRegionOnly `
    -OutputDirectory '.\SEA-regional-only'

# See canonical region names (not display names such as "Southeast Asia").
az account list-locations --query "[].{Name:name,DisplayName:displayName}" -o table
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

The collector does this independently for **each subscription** and for resource
groups as a separate dataset:

1. Query `summarize ResourceCount=count()` for the collection scope.
2. Query resources ordered by their unique `id`, retaining `id` in the projection.
3. Request **`--first 1000`**.
4. Pass the response's `skip_token` to **`--skip-token`** on the next request.
5. Continue until there is no continuation token, even if the initial count has
   already been reached. This avoids stopping early when the inventory changes.
6. Check duplicate IDs, repeated tokens, explicit truncation, the downloaded
   unique count, and a second count query.

For example, 25,342 resources in one subscription require 26 resource pages, not
one query capped at 1,000. The two count queries are additional requests.
The CLI spelling is `--first`; PowerShell `Search-AzGraph` uses `-First`.
This implementation does not need `Search-AzGraph`.

Continuation tokens are safer than repeatedly issuing `--first 1000` alone,
which would return the first page again. There are no `take`, `limit`, or
`mv-expand` operators in the resource paging query. Dependency extraction happens
locally, avoiding expansion limits and join-related fan-out in ARG.

Preflight adds one live ARG count request per selected subscription before
collection starts. Transient throttling, timeout and selected server failures receive bounded
exponential backoff with jitter (six retries by default). Authorization errors
fail that subscription rather than being retried indefinitely.

ARG is eventually consistent and does not offer a transactionally frozen
snapshot for this scan. Equal before/after counts cannot detect replacement of
one resource by another or all concurrent property changes. Collect in a stable
window and rerun discrepancies.

## Workbook

| Worksheet | Contents |
|---|---|
| Summary | Snapshot counts by subscription, resource type and impact scope; counts with no tags and no exposed/applicable SKU; formula totals. |
| Inventory | Impacted resource IDs/names/types, groups, subscriptions and tenants, declared/effective regions, normalized SKU, zones, state, identity metadata, tags, selected service details and migration review prompts. |
| Dependencies | Configured outbound and inbound ARM references touching an impacted resource; cross-region/subscription flags, evidence paths and resolution status. |
| DependencyImpact | Per-impacted-resource inbound/outbound reference counts, cross-region references and unresolved references. Counts are evidence rows, not distinct applications or traffic flows. |
| GlobalContext | Global or unspecified-location resources for a separate review; not automatically counted as impacted. |
| ResourceGroups | Group metadata location and tags. Group location is not workload deployment location. |
| ScopeAudit | Per-subscription, per-dataset expected/downloaded counts, page counts and reconciliation status. |
| Coverage | Interpretation, scope, caveats and migration checklist. |
| Errors | Explicit subscription failures. Empty on a fully successful run. |

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
inferences. It resolves references after every subscription has been scanned.
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

Run from PowerShell 7 in the folder containing the script and module:

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
| `Required module file missing` | Only the entry script was copied, or the companion file was renamed | Place `RegionalInventory.psm1` beside `Export-AzureRegionalInventory.ps1`. Do not install an unrelated Gallery module with a similar name. |
| Script execution is blocked or the file is not digitally signed | Execution policy, downloaded-file marking, or organizational application control | Inspect `Get-ExecutionPolicy -List`. Use your approved signing/unblocking process or contact your administrator. Do not bypass organizational policy or disable security controls. |
| `Azure CLI is not on PATH`, or `az` is not recognized | CLI is missing or the terminal has an old PATH | Install Azure CLI using your approved method, reopen the terminal and check `Get-Command az`. |
| `Azure CLI 2.22.0+ is required` | CLI is too old or its version response is invalid | Check `az version`; update using your approved CLI installation method. |
| `Missing resource-graph extension`, or `graph` is not recognized | Extension is not installed or is incompatible | Run `az extension add --name resource-graph`. If already installed but incompatible, run `az extension update --name resource-graph`, subject to your organization's change policy. |
| `Azure CLI session could not be refreshed`, `Please run az login`, or an AADSTS/token-expiration error | Missing or expired CLI authentication, wrong tenant, or Conditional Access requirements | Run `az login --tenant '<tenant-guid>'`, complete the required authentication, then rerun preflight. An existing `Connect-AzAccount` session does not sign in the CLI. |
| Browser-based login cannot complete | Browser access or authentication flow is restricted | If your organization permits device-code authentication, use `-Login -UseDeviceCode`. If Conditional Access blocks it, use an approved authentication flow with your administrator. |
| `UseDeviceCode requires Login` | Device-code option was supplied without requesting sign-in | Supply both switches: `-Login -UseDeviceCode`. |
| `Subscription not in CLI account list` or `No subscriptions in the requested scope` | Incorrect tenant/ID, missing subscription visibility, or stale authentication | Sign in to the correct tenant, check `az account list --all --refresh -o table`, and verify the requested subscription GUIDs and Reader access. Do not remove required subscriptions just to obtain a passing run. |
| Subscription is `Disabled` or otherwise not `Enabled` | Subscription state does not permit this collection | Ask the subscription administrator to investigate. Exclude it only if the inventory owner explicitly accepts the reduced scope and records the gap. |

### Azure query and pagination errors

| Error or symptom | Likely cause | Fix |
|---|---|---|
| `Live ARG preflight failed`, `AuthorizationFailed` or HTTP 403 | Token/tenant mismatch, insufficient RBAC, or another access restriction | Read the original error included in the message. Refresh login for the subscription's tenant and have the administrator verify Reader access at the intended scope. After a role change, allow propagation and rerun preflight. |
| SSL/certificate, proxy, DNS, connection or timeout errors | Corporate proxy/VPN/firewall or certificate trust is preventing CLI access | Confirm access to the configured Azure cloud through your approved network. Ask IT to verify proxy configuration and trusted enterprise certificates. Do not disable TLS verification. |
| `429`, `TooManyRequests`, or transient 502/503/504 warnings | Azure throttling or a transient service failure | Allow the built-in bounded retries to finish. If exhausted, reduce competing inventory jobs, wait and rerun with a new output path. Splitting subscription batches can reduce competing load; keep records of the full intended scope. |
| `Unknown source region` / `Unknown target region` | A display name or misspelled region was supplied | Use the canonical `name` from `az account list-locations -o table`, such as `southeastasia`, not `Southeast Asia`. |
| `SourceRegion and TargetRegion must be different` | The same region was supplied twice | Correct the destination, or omit `-TargetRegion` when only gathering inventory. |
| `ARG count mismatch` or `Duplicate resource ID while paging` | Resources changed during collection, ARG indexing changed, or paging was inconsistent | Rerun during a quieter window using a new output folder. Keep the error/count evidence if it recurs. Do not disable reconciliation to force a successful result. |
| Repeated continuation token, empty page with a token, or truncation without a token | Unexpected ARG/extension response or a modified query that is not pageable | Update an incompatible extension through your approved process and retry. If you changed the query, retain scalar `id`, stable ordering and token handling; remove `take`, `limit` or `sample` operators. Never treat the first 1,000 rows as the complete result. |
| Invalid ARG count response, missing `data`, or CLI returned no JSON | Unexpected CLI output, incompatible tooling or a custom query change | Check CLI/extension versions and reproduce the small count query below. Preserve the original error for support; do not replace failed responses with empty inventories. |

To isolate a live query problem without running the entire collector:

```powershell
az graph query `
    --subscription '<subscription-guid>' `
    --subscriptions '<subscription-guid>' `
    --first 1 `
    -q 'Resources | summarize ResourceCount=count()' `
    --output json
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
| A cell ends with `[TRUNCATED: see JSONL companion]` | The selected metadata exceeded Excel's cell text limit | Read the corresponding resource/row in the JSONL companion. This is an explicit display limit, not evidence that the entire collection stopped. |

### Incomplete or unexpected results

- **`PARTIAL inventory saved`:** inspect the `Errors` and `ScopeAudit` worksheets,
  then fix the reported subscription failures. Some rows may already have been
  saved from a failed subscription. Do not present that workbook as a complete
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

Record the failing stage, sanitized error message, UTC time, CLI/extension/runtime
versions, and whether the issue occurs in `-CheckOnly`. For collection failures,
also inspect `Run.json`, `Preflight.json` when present, and the relevant
`ScopeAudit`/`Errors` rows.

```powershell
$PSVersionTable.PSVersion
[Environment]::Is64BitProcess
Get-Location
Get-Command az | Select-Object Name, Source
az version
az extension show --name resource-graph --query '{name:name,version:version}' -o json
```

Sanitize subscription/tenant IDs, resource names, tags, paths and other client
metadata before sharing outside the approved support channel. Never share
access tokens, credentials, secret values or unreviewed CLI debug logs. The
diagnostic commands above do not request token values.

## Offline verification

```powershell
.\Test-RegionalInventory.ps1 -Scale
```

Tests use mocked CLI responses and make **no Azure connection**. They cover
preflight-only mode, CLI version/extension/session checks, live-query failures,
disabled or missing subscriptions, invalid regions and output paths, plus a
100,001-resource subscription, >1,000-resource orchestration, continuation tokens,
empty/exact-page scopes, duplicates, truncation, count mismatches, retry behavior,
service-specific SKUs, late cross-subscription plan resolution, secondary-region
storage, inherited child location, inbound/cross-region references, partial
failures, worksheet splitting, XML validity, formula injection prevention,
text limits, and cached formula totals.
If ImportExcel is already installed, tests additionally open the workbook with
its independent EPPlus reader and recalculate summary totals. ImportExcel is not
a production dependency. No dependencies are installed by the tests.

## References

- [ARG large datasets, page size and continuation tokens](https://learn.microsoft.com/en-us/azure/governance/resource-graph/concepts/work-with-data)
- [Azure CLI graph query parameters](https://learn.microsoft.com/en-us/cli/azure/graph#az-graph-query)
- [ARG supported resource types and tables](https://learn.microsoft.com/en-us/azure/governance/resource-graph/reference/supported-tables-resources)
