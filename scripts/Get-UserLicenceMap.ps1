<#
.SYNOPSIS
    Maps per-user licence assignments across the tenant.

.DESCRIPTION
    Retrieves all users with their assigned licences, sign-in activity, and profile
    data via Microsoft Graph. Produces a flat CSV with one row per user-licence
    combination for downstream waste analysis.

    Requires an active Microsoft Graph connection with User.Read.All permission.
    For sign-in activity, AuditLog.Read.All is also needed.

.PARAMETER OutputPath
    Directory to write the CSV output. Defaults to ./output.

.PARAMETER ConfigPath
    Path to the config directory containing sku-names.json.
    Defaults to ../config relative to the script location.

.EXAMPLE
    Connect-MgGraph -Scopes "User.Read.All","Organization.Read.All"
    ./Get-UserLicenceMap.ps1 -OutputPath ./output
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

# --- Load SKU data from inventory step ---
$skuDataFile = Join-Path $OutputPath "sku-data.json"
if (-not (Test-Path $skuDataFile)) {
    Write-Error "SKU data not found at $skuDataFile — run Get-LicenceInventory.ps1 first"
    return
}
$skus = Get-Content $skuDataFile -Raw | ConvertFrom-Json

# --- Load friendly names ---
$skuNamesFile = Join-Path $ConfigPath "sku-names.json"
if (Test-Path $skuNamesFile) {
    $skuNames = Get-Content $skuNamesFile -Raw | ConvertFrom-Json -AsHashtable
} else {
    $skuNames = @{}
}

# --- Ensure output directory ---
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# --- Pull all users ---
Write-Host "Retrieving all users with licence and sign-in data..." -ForegroundColor Cyan
Write-Host "(This may take a few minutes for large tenants)"

$users = Get-MgUser -All -Property @(
    "Id", "DisplayName", "UserPrincipalName", "AccountEnabled",
    "UserType", "AssignedLicenses", "SignInActivity", "Department", "JobTitle"
) -ConsistencyLevel eventual -CountVariable userCount

Write-Host "Retrieved $($users.Count) users" -ForegroundColor Cyan

# --- Build per-user licence map ---
$userLicences = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($user in $users) {
    foreach ($licence in $user.AssignedLicenses) {
        $skuMatch = $skus | Where-Object { $_.SkuId -eq $licence.SkuId }
        $skuName = if ($skuMatch) { $skuMatch.SkuPartNumber } else { $licence.SkuId }

        $userLicences.Add([PSCustomObject]@{
            UPN           = $user.UserPrincipalName
            DisplayName   = $user.DisplayName
            Enabled       = $user.AccountEnabled
            UserType      = $user.UserType
            Department    = $user.Department
            JobTitle      = $user.JobTitle
            SKU           = $skuName
            FriendlyName  = if ($null -ne $skuNames[$skuName]) { $skuNames[$skuName] } else { $skuName }
            DisabledPlans = ($licence.DisabledPlans | Measure-Object).Count
            LastSignIn    = $user.SignInActivity.LastSignInDateTime
        })
    }
}

# --- Output ---
$outputFile = Join-Path $OutputPath "user-licence-map.csv"
$userLicences | Export-Csv -Path $outputFile -NoTypeInformation
Write-Host "`nUser-licence map written to $outputFile" -ForegroundColor Green

# --- Summary ---
$uniqueUsers = ($userLicences | Select-Object -Unique UPN | Measure-Object).Count
$totalAssignments = $userLicences.Count
$disabledWithLicences = ($userLicences | Where-Object { $_.Enabled -eq $false } |
    Select-Object -Unique UPN | Measure-Object).Count

Write-Host "Unique users with licences: $uniqueUsers"
Write-Host "Total licence assignments: $totalAssignments"
Write-Host "Disabled accounts with licences: $disabledWithLicences"
