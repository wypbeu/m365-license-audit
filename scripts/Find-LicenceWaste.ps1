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
Write-Host "Pricing file: $skuPricingFile" -ForegroundColor DarkGray
if (Test-Path $skuPricingFile) {
    $skuPricing = Get-Content $skuPricingFile -Raw | ConvertFrom-Json -AsHashtable
    @($skuPricing.Keys) | Where-Object { $_ -like "_*" } | ForEach-Object { [void]$skuPricing.Remove($_) }
    Write-Host "Pricing loaded: $($skuPricing.Count) SKU entries" -ForegroundColor DarkGray
} else {
    Write-Warning "SKU pricing config not found at $skuPricingFile — cost estimates will be zero"
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

    # Microsoft has occasionally emitted duplicate column headers in this report
    # (Import-Csv then throws "The member X is already present"). Sanitise the
    # header row before parsing by suffixing repeats with _1, _2, ...
    $rawLines = Get-Content $copilotPath
    if ($rawLines.Count -gt 0) {
        $headerCols = $rawLines[0] -split ','
        $seen = @{}
        $clean = foreach ($h in $headerCols) {
            $key = $h.Trim('"')
            if ($seen.ContainsKey($key)) {
                $seen[$key]++
                "$key`_$($seen[$key])"
            } else {
                $seen[$key] = 0
                $key
            }
        }
        $rawLines[0] = ($clean | ForEach-Object { '"' + $_ + '"' }) -join ','
        $rawLines | Set-Content -Path $copilotPath -Encoding UTF8
    }

    # Report can return multiple rows per user (one per active day); deduplicate
    $copilotUsage = @(Import-Csv $copilotPath |
        Sort-Object 'User Principal Name', 'Last Activity Date' -Descending |
        Sort-Object 'User Principal Name' -Unique)
    $hasCopilotData = $true
    Write-Host "  Copilot usage rows: $($copilotUsage.Count)" -ForegroundColor DarkGray
} catch {
    Write-Warning "Copilot usage report not available (beta endpoint): $_"
    $copilotUsage = @()
    $hasCopilotData = $false
}

$cutoffDate = (Get-Date).AddDays(-$InactiveDays)
$allWaste = [System.Collections.Generic.List[PSCustomObject]]::new()

# Helper to check if a SKU is free/zero-cost.
# Defaults to "paid" (false) when pricing is unknown so the analysis still
# produces records if sku-pricing.json is missing — otherwise the script
# silently filters everything out and writes a 0-byte waste CSV.
function Test-FreeSku {
    param([string]$SkuName)
    if ($SkuName -match "FREE|FLOW_FREE|POWER_BI_STANDARD|_viral|_vTrial|TRIAL") { return $true }
    if ($skuPricing.ContainsKey($SkuName)) {
        return ([double]$skuPricing[$SkuName] -eq 0)
    }
    return $false  # unknown SKU — treat as paid, include in analysis
}

# --- Pattern 1: Disabled accounts with paid licences ---
Write-Host "`nAnalysing waste pattern: Disabled accounts..." -ForegroundColor Yellow
try {
    $disabled = @($userLicences | Where-Object {
        $_.Enabled -eq "False" -and -not (Test-FreeSku $_.SKU)
    })
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
} catch {
    Write-Warning "Pattern 1 (Disabled) failed: $($_.Exception.Message)"
    Write-Warning "  at $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
}

# --- Pattern 2: Service accounts on premium licences ---
Write-Host "Analysing waste pattern: Service accounts..." -ForegroundColor Yellow
try {
    $serviceAccounts = @($userLicences | Where-Object {
        $_.Enabled -eq "True" -and
        $_.SKU -match "ENTERPRISE|SPE_E" -and
        (
            $_.UPN -match $ServiceAccountPattern -or
            [string]::IsNullOrWhiteSpace($_.LastSignIn) -or
            (-not [string]::IsNullOrWhiteSpace($_.LastSignIn) -and [datetime]$_.LastSignIn -lt (Get-Date).AddDays(-180))
        )
    })
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
} catch {
    Write-Warning "Pattern 2 (Service accounts) failed: $($_.Exception.Message)"
    Write-Warning "  at $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
}

# --- Pattern 3: Guests with paid licences ---
Write-Host "Analysing waste pattern: Guest users..." -ForegroundColor Yellow
try {
    $guests = @($userLicences | Where-Object {
        $_.UserType -eq "Guest" -and -not (Test-FreeSku $_.SKU)
    })
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
} catch {
    Write-Warning "Pattern 3 (Guests) failed: $($_.Exception.Message)"
    Write-Warning "  at $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
}

# --- Pattern 4: Inactive users ---
Write-Host "Analysing waste pattern: Inactive users ($InactiveDays+ days)..." -ForegroundColor Yellow
try {
    $activeUsers = @($userLicences | Where-Object {
        $_.Enabled -eq "True" -and
        $_.UserType -ne "Guest" -and
        -not ($_.UPN -match $ServiceAccountPattern) -and
        -not (Test-FreeSku $_.SKU)
    })

    # Index usage by UPN once to avoid O(n*m) scans on large tenants
    $usageByUpn = @{}
    foreach ($row in $usage) {
        $k = $row.'User Principal Name'
        if (-not [string]::IsNullOrWhiteSpace($k)) { $usageByUpn[$k] = $row }
    }

    foreach ($record in $activeUsers) {
        if (-not $usageByUpn.ContainsKey($record.UPN)) { continue }
        $activity = $usageByUpn[$record.UPN]

        $dateStrings = @(
            $activity.'Exchange Last Activity Date'
            $activity.'Teams Last Activity Date'
            $activity.'SharePoint Last Activity Date'
            $activity.'OneDrive Last Activity Date'
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        $lastActivity = $null
        foreach ($s in $dateStrings) {
            try {
                $d = [datetime]$s
                if ($null -eq $lastActivity -or $d -gt $lastActivity) { $lastActivity = $d }
            } catch {}
        }

        if ($null -eq $lastActivity -or $lastActivity -lt $cutoffDate) {
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
    $inactiveCount = @($allWaste | Where-Object { $_.WasteCategory -eq "Inactive User" }).Count
    Write-Host "  Found $inactiveCount licence assignments on inactive users"
} catch {
    Write-Warning "Pattern 4 (Inactive) failed: $($_.Exception.Message)"
    Write-Warning "  at $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
}

# --- Pattern 5: Copilot with no activity ---
if ($hasCopilotData) {
    Write-Host "Analysing waste pattern: Copilot unused..." -ForegroundColor Yellow
    try {
        $copilotLicensed = @($userLicences | Where-Object { $_.SKU -match "Copilot" })
        $copilotActiveUPNs = @($copilotUsage | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.'Last Activity Date')
        } | ForEach-Object { $_.'User Principal Name' })

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
        $copilotWaste = @($allWaste | Where-Object { $_.WasteCategory -eq "Copilot Unused" }).Count
        Write-Host "  Found $copilotWaste Copilot licences with no activity"
    } catch {
        Write-Warning "Pattern 5 (Copilot) failed: $($_.Exception.Message)"
        Write-Warning "  at $($_.InvocationInfo.ScriptLineNumber): $($_.InvocationInfo.Line.Trim())"
    }
}

# --- Output ---
$outputFile = Join-Path $OutputPath "licence-waste.csv"
$allWaste | Export-Csv -Path $outputFile -NoTypeInformation
Write-Host "`nAll waste records written to $outputFile" -ForegroundColor Green
Write-Host "Total waste records: $($allWaste.Count)"
