<#
.SYNOPSIS
    Identifies licence waste patterns by cross-referencing assignments against usage.

.DESCRIPTION
    Reads the user-licence map from Get-UserLicenceMap.ps1 and pulls M365 usage
    reports via Graph API. Identifies five waste patterns:
      1. Disabled accounts with paid licences
      2. Service accounts on premium licences
      3. Guest users with paid licences
      4. Inactive users (no workload activity in 90+ days)
      5. Copilot licences with no adoption

    Requires an active Microsoft Graph connection with Reports.Read.All permission.
    Usage reports must have user-level detail enabled (not obfuscated).

.PARAMETER OutputPath
    Directory containing user-licence-map.csv and where waste CSVs are written.
    Defaults to ./output.

.PARAMETER ConfigPath
    Path to the config directory containing sku-pricing.json.
    Defaults to ../config relative to the script location.

.PARAMETER InactiveDays
    Number of days without activity to classify a user as inactive. Defaults to 90.

.PARAMETER ServiceAccountPattern
    Regex pattern to identify service account UPNs. Defaults to common prefixes.

.EXAMPLE
    ./Find-LicenceWaste.ps1 -OutputPath ./output -InactiveDays 120
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = (Join-Path $PSScriptRoot ".." "output"),

    [Parameter()]
    [string]$ConfigPath = (Join-Path $PSScriptRoot ".." "config"),

    [Parameter()]
    [int]$InactiveDays = 90,

    [Parameter()]
    [string]$ServiceAccountPattern = "^(svc|service|admin|noreply|do-not-reply|mailbox|room|shared)"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Load user-licence map ---
$mapFile = Join-Path $OutputPath "user-licence-map.csv"
if (-not (Test-Path $mapFile)) {
    Write-Error "User-licence map not found at $mapFile — run Get-UserLicenceMap.ps1 first"
    return
}
$userLicences = Import-Csv $mapFile

# --- Load pricing ---
$skuPricingFile = Join-Path $ConfigPath "sku-pricing.json"
if (Test-Path $skuPricingFile) {
    $skuPricing = Get-Content $skuPricingFile -Raw | ConvertFrom-Json -AsHashtable
    $skuPricing.Remove("_comment")
} else {
    Write-Warning "SKU pricing config not found — cost estimates will be zero"
    $skuPricing = @{}
}

# --- Pull usage report ---
Write-Host "Pulling M365 usage report (180-day period)..." -ForegroundColor Cyan
$usagePath = Join-Path $env:TEMP "M365Usage_$(Get-Date -Format yyyyMMdd).csv"
$uri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D180')"
Invoke-MgGraphRequest -Uri $uri -OutputFilePath $usagePath
$usage = Import-Csv $usagePath

# --- Pull Copilot usage (beta) ---
Write-Host "Pulling Copilot usage report..." -ForegroundColor Cyan
$copilotPath = Join-Path $env:TEMP "CopilotUsage_$(Get-Date -Format yyyyMMdd).csv"
try {
    $copilotUri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='D180')"
    Invoke-MgGraphRequest -Uri $copilotUri -OutputFilePath $copilotPath
    $copilotUsage = Import-Csv $copilotPath
    $hasCopilotData = $true
} catch {
    Write-Warning "Copilot usage report not available (beta endpoint): $_"
    $copilotUsage = @()
    $hasCopilotData = $false
}

$cutoffDate = (Get-Date).AddDays(-$InactiveDays)
$allWaste = [System.Collections.Generic.List[PSCustomObject]]::new()

# Helper to check if a SKU is free/zero-cost
function Test-FreeSku {
    param([string]$SkuName)
    $cost = $skuPricing[$SkuName]
    return (-not $cost -or $cost -eq 0 -or $SkuName -match "FREE|FLOW_FREE|POWER_BI_STANDARD")
}

# --- Pattern 1: Disabled accounts with paid licences ---
Write-Host "`nAnalysing waste pattern: Disabled accounts..." -ForegroundColor Yellow
$disabled = $userLicences | Where-Object {
    $_.Enabled -eq "False" -and -not (Test-FreeSku $_.SKU)
}
foreach ($record in $disabled) {
    $allWaste.Add([PSCustomObject]@{
        UPN              = $record.UPN
        DisplayName      = $record.DisplayName
        Department       = $record.Department
        SKU              = $record.SKU
        FriendlyName     = $record.FriendlyName
        WasteCategory    = "Disabled Account"
        LastSignIn       = $record.LastSignIn
        MonthlyCost      = $skuPricing[$record.SKU] ?? 0
    })
}
Write-Host "  Found $($disabled.Count) licence assignments on disabled accounts"

# --- Pattern 2: Service accounts on premium licences ---
Write-Host "Analysing waste pattern: Service accounts..." -ForegroundColor Yellow
$serviceAccounts = $userLicences | Where-Object {
    $_.Enabled -eq "True" -and
    $_.SKU -match "ENTERPRISE|SPE_E" -and
    (
        $_.UPN -match $ServiceAccountPattern -or
        $_.LastSignIn -eq "" -or
        ($_.LastSignIn -ne "" -and [datetime]$_.LastSignIn -lt (Get-Date).AddDays(-180))
    )
}
foreach ($record in $serviceAccounts) {
    $allWaste.Add([PSCustomObject]@{
        UPN              = $record.UPN
        DisplayName      = $record.DisplayName
        Department       = $record.Department
        SKU              = $record.SKU
        FriendlyName     = $record.FriendlyName
        WasteCategory    = "Service Account"
        LastSignIn       = $record.LastSignIn
        MonthlyCost      = $skuPricing[$record.SKU] ?? 0
    })
}
Write-Host "  Found $($serviceAccounts.Count) potential service accounts on premium licences"

# --- Pattern 3: Guests with paid licences ---
Write-Host "Analysing waste pattern: Guest users..." -ForegroundColor Yellow
$guests = $userLicences | Where-Object {
    $_.UserType -eq "Guest" -and -not (Test-FreeSku $_.SKU)
}
foreach ($record in $guests) {
    $allWaste.Add([PSCustomObject]@{
        UPN              = $record.UPN
        DisplayName      = $record.DisplayName
        Department       = $record.Department
        SKU              = $record.SKU
        FriendlyName     = $record.FriendlyName
        WasteCategory    = "Guest"
        LastSignIn       = $record.LastSignIn
        MonthlyCost      = $skuPricing[$record.SKU] ?? 0
    })
}
Write-Host "  Found $($guests.Count) guest users with paid licences"

# --- Pattern 4: Inactive users ---
Write-Host "Analysing waste pattern: Inactive users ($InactiveDays+ days)..." -ForegroundColor Yellow
$activeUsers = $userLicences | Where-Object {
    $_.Enabled -eq "True" -and
    $_.UserType -ne "Guest" -and
    -not ($_.UPN -match $ServiceAccountPattern) -and
    -not (Test-FreeSku $_.SKU)
}

foreach ($record in $activeUsers) {
    $activity = $usage | Where-Object { $_.'User Principal Name' -eq $record.UPN }
    if (-not $activity) { continue }

    $lastActivity = @(
        $activity.'Exchange Last Activity Date'
        $activity.'Teams Last Activity Date'
        $activity.'SharePoint Last Activity Date'
        $activity.'OneDrive Last Activity Date'
    ) | Where-Object { $_ -ne "" } |
        ForEach-Object { [datetime]$_ } |
        Sort-Object -Descending |
        Select-Object -First 1

    if (-not $lastActivity -or $lastActivity -lt $cutoffDate) {
        $allWaste.Add([PSCustomObject]@{
            UPN              = $record.UPN
            DisplayName      = $record.DisplayName
            Department       = $record.Department
            SKU              = $record.SKU
            FriendlyName     = $record.FriendlyName
            WasteCategory    = "Inactive User"
            LastSignIn       = $record.LastSignIn
            MonthlyCost      = $skuPricing[$record.SKU] ?? 0
        })
    }
}
$inactiveCount = ($allWaste | Where-Object { $_.WasteCategory -eq "Inactive User" } | Measure-Object).Count
Write-Host "  Found $inactiveCount licence assignments on inactive users"

# --- Pattern 5: Copilot with no activity ---
if ($hasCopilotData) {
    Write-Host "Analysing waste pattern: Copilot unused..." -ForegroundColor Yellow
    $copilotLicensed = $userLicences | Where-Object { $_.SKU -match "Copilot" }
    $copilotActiveUPNs = ($copilotUsage | Where-Object { $_.'Last Activity Date' -ne "" }).'User Principal Name'

    foreach ($record in $copilotLicensed) {
        if ($record.UPN -notin $copilotActiveUPNs) {
            $allWaste.Add([PSCustomObject]@{
                UPN              = $record.UPN
                DisplayName      = $record.DisplayName
                Department       = $record.Department
                SKU              = $record.SKU
                FriendlyName     = $record.FriendlyName
                WasteCategory    = "Copilot Unused"
                LastSignIn       = $record.LastSignIn
                MonthlyCost      = $skuPricing[$record.SKU] ?? 0
            })
        }
    }
    $copilotWaste = ($allWaste | Where-Object { $_.WasteCategory -eq "Copilot Unused" } | Measure-Object).Count
    Write-Host "  Found $copilotWaste Copilot licences with no activity"
}

# --- Output ---
$outputFile = Join-Path $OutputPath "licence-waste.csv"
$allWaste | Export-Csv -Path $outputFile -NoTypeInformation
Write-Host "`nAll waste records written to $outputFile" -ForegroundColor Green
Write-Host "Total waste records: $($allWaste.Count)"
