<#
.SYNOPSIS
    Pulls the licence inventory for the connected M365 tenant.

.DESCRIPTION
    Retrieves all subscribed SKUs via Microsoft Graph and produces a summary
    showing purchased, assigned, and available counts with utilisation percentage.
    Resolves cryptic SKU part numbers to friendly names using the config file.

    Requires an active Microsoft Graph connection with Organization.Read.All permission.

.PARAMETER OutputPath
    Directory to write the CSV output. Defaults to ./output.

.PARAMETER ConfigPath
    Path to the config directory containing sku-names.json and sku-pricing.json.
    Defaults to ../config relative to the script location.

.EXAMPLE
    Connect-MgGraph -Scopes "Organization.Read.All"
    ./Get-LicenceInventory.ps1 -OutputPath ./output
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = (Join-Path $PSScriptRoot ".." "output"),

    [Parameter()]
    [string]$ConfigPath = (Join-Path $PSScriptRoot ".." "config")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Load config ---
$skuNamesFile = Join-Path $ConfigPath "sku-names.json"
$skuPricingFile = Join-Path $ConfigPath "sku-pricing.json"

if (-not (Test-Path $skuNamesFile)) {
    Write-Warning "SKU names config not found at $skuNamesFile — using raw SKU part numbers"
    $skuNames = @{}
} else {
    $skuNames = Get-Content $skuNamesFile -Raw | ConvertFrom-Json -AsHashtable
}

if (-not (Test-Path $skuPricingFile)) {
    Write-Warning "SKU pricing config not found at $skuPricingFile — cost columns will be zero"
    $skuPricing = @{}
} else {
    $skuPricing = Get-Content $skuPricingFile -Raw | ConvertFrom-Json -AsHashtable
    $skuPricing.Remove("_comment")
}

# --- Ensure output directory ---
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# --- Pull subscribed SKUs ---
Write-Host "Retrieving subscribed SKUs..." -ForegroundColor Cyan
$skus = Get-MgSubscribedSku -All

$licenceSummary = $skus |
    Where-Object { $_.AppliesTo -eq "User" } |
    Select-Object @{
        Name = "SKUPartNumber"; Expression = { $_.SkuPartNumber }
    }, @{
        Name = "FriendlyName"; Expression = { if ($null -ne $skuNames[$_.SkuPartNumber]) { $skuNames[$_.SkuPartNumber] } else { $_.SkuPartNumber } }
    }, @{
        Name = "Purchased"; Expression = { $_.PrepaidUnits.Enabled }
    }, @{
        Name = "Assigned"; Expression = { $_.ConsumedUnits }
    }, @{
        Name = "Available"; Expression = { $_.PrepaidUnits.Enabled - $_.ConsumedUnits }
    }, @{
        Name = "Utilisation"; Expression = {
            if ($_.PrepaidUnits.Enabled -gt 0) {
                [math]::Round(($_.ConsumedUnits / $_.PrepaidUnits.Enabled) * 100, 1)
            } else { 0 }
        }
    }, @{
        Name = "EstimatedMonthlySpend"; Expression = {
            $cost = $skuPricing[$_.SkuPartNumber]
            if ($cost) { [math]::Round($_.ConsumedUnits * $cost, 2) } else { 0 }
        }
    }

# --- Output ---
$licenceSummary | Sort-Object SKUPartNumber | Format-Table -AutoSize

$outputFile = Join-Path $OutputPath "licence-inventory.csv"
$licenceSummary | Sort-Object SKUPartNumber | Export-Csv -Path $outputFile -NoTypeInformation
Write-Host "Inventory written to $outputFile" -ForegroundColor Green

# --- Summary stats ---
$totalMonthly = ($licenceSummary | Measure-Object -Property EstimatedMonthlySpend -Sum).Sum
$totalPurchased = ($licenceSummary | Measure-Object -Property Purchased -Sum).Sum
$totalAssigned = ($licenceSummary | Measure-Object -Property Assigned -Sum).Sum

Write-Host "`nTotal SKUs: $($licenceSummary.Count)"
Write-Host "Total licences purchased: $totalPurchased"
Write-Host "Total licences assigned: $totalAssigned"
Write-Host "Estimated monthly spend: £$('{0:N0}' -f $totalMonthly)"

# --- Export SKU data for downstream scripts ---
$skuDataFile = Join-Path $OutputPath "sku-data.json"
$skus | ConvertTo-Json -Depth 5 | Set-Content $skuDataFile
Write-Host "Raw SKU data written to $skuDataFile (used by downstream scripts)"
