#Requires -Version 7.2
<#
.SYNOPSIS
Creates a read-only, paginated Azure regional migration inventory and Excel workbook.
.EXAMPLE
.\Export-AzureRegionalInventory.ps1 -SourceRegion southeastasia -TargetRegion eastasia -Login
.EXAMPLE
.\Export-AzureRegionalInventory.ps1 -SourceRegion southeastasia -SubscriptionId '<subscription-guid>' -Verbose
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidatePattern('^[a-z0-9]+$')][string]$SourceRegion,
    [ValidatePattern('^[a-z0-9]*$')][string]$TargetRegion = '',
    [string[]]$SubscriptionId,
    [string]$TenantId,
    [string]$OutputDirectory = (Join-Path $PWD ("RegionalInventory-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))),
    [switch]$Login,
    [switch]$UseDeviceCode,
    [switch]$CheckOnly,
    [ValidateSet('AllAccessibleRegions', 'SourceRegionOnly')][string]$DependencyScope = 'AllAccessibleRegions',
    [ValidateRange(1, 1048575)][int]$RowsPerSheet = 250000
)
$ErrorActionPreference = 'Stop'
$modulePath = Join-Path $PSScriptRoot 'RegionalInventory.psm1'
if (-not (Test-Path -LiteralPath $modulePath -PathType Leaf)) {
    throw "Required module file missing: $modulePath. Keep RegionalInventory.psm1 beside this script; no Gallery modules are required."
}
Import-Module $modulePath -Force
Invoke-RegionalInventory @PSBoundParameters
