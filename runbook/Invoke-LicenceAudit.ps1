<#
.SYNOPSIS
    Azure Automation runbook that performs a full M365 licence audit and writes
    results to SharePoint lists for Power BI consumption.

.DESCRIPTION
    Consolidates the interactive audit scripts into a single unattended runbook.
    Authenticates via System-Assigned Managed Identity, pulls licence inventory
    and per-user assignments, identifies waste patterns, and writes records to
    two SharePoint lists: LicenceAuditResults and LicenceInventorySummary.

    Schedule: Weekly (recommended Sunday 02:00 UTC)
    Identity: System-Assigned Managed Identity on the Automation Account

    Required Graph application permissions on the Managed Identity:
      - Organization.Read.All
      - User.Read.All
      - Reports.Read.All
      - Directory.Read.All
      - Sites.Manage.All

.NOTES
    Grant permissions to the Managed Identity using:
    New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId <MI-object-id> ...

    See the companion blog post for full setup instructions:
    https://sbd.org.uk/blog/m365-licensing-audit
#>

# ============================================================================
# CONFIGURATION — Update these values for your environment
# ============================================================================

$siteId        = "contoso.sharepoint.com,<site-guid>,<web-guid>"
$auditListId   = "<LicenceAuditResults-list-guid>"
$summaryListId = "<LicenceInventorySummary-list-guid>"

# Monthly per-user cost in GBP — update to match your EA/CSP pricing
$skuMonthlyCost = @{
    "ENTERPRISEPREMIUM"     = 49.20  # Microsoft 365 E5
    "ENTERPRISEPACK"        = 28.40  # Microsoft 365 E3
    "SPE_E5"                = 49.20  # Microsoft 365 E5 (unified)
    "SPE_E3"                = 28.40  # Microsoft 365 E3 (unified)
    "SPE_F1"                = 7.50   # Microsoft 365 F3
    "EMSPREMIUM"            = 12.30  # Enterprise Mobility + Security E5
    "EMS"                   = 7.30   # Enterprise Mobility + Security E3
    "Microsoft_365_Copilot" = 24.00  # Microsoft 365 Copilot
    "TEAMS_PREMIUM"         = 8.40   # Teams Premium
    "INTUNE_A"              = 6.80   # Intune Plan 1
    "INTUNE_P1"             = 6.80   # Intune Plan 1 (standalone)
}

$skuFriendlyNames = @{
    "ENTERPRISEPREMIUM"     = "Microsoft 365 E5"
    "ENTERPRISEPACK"        = "Microsoft 365 E3"
    "SPE_E5"                = "Microsoft 365 E5 (unified)"
    "SPE_E3"                = "Microsoft 365 E3 (unified)"
    "SPE_F1"                = "Microsoft 365 F3"
    "EMSPREMIUM"            = "Enterprise Mobility + Security E5"
    "EMS"                   = "Enterprise Mobility + Security E3"
    "Microsoft_365_Copilot" = "Microsoft 365 Copilot"
    "TEAMS_PREMIUM"         = "Teams Premium"
}

# UPN prefix patterns that identify service/shared/room accounts
$serviceAccountPattern = "^(svc|service|admin|noreply|do-not-reply|mailbox|room|shared)"

# Days of inactivity before classifying a user as inactive
$inactiveDays = 90

# ============================================================================
# EXECUTION
# ============================================================================

$ErrorActionPreference = "Stop"
$auditDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
$cutoffDate = (Get-Date).AddDays(-$inactiveDays)

# --- Authenticate ---
Write-Output "Connecting to Microsoft Graph via Managed Identity..."
Connect-MgGraph -Identity

# --- Step 1: Licence Inventory ---
Write-Output "Retrieving subscribed SKUs..."
$skus = Get-MgSubscribedSku -All

foreach ($sku in ($skus | Where-Object { $_.AppliesTo -eq "User" })) {
    $utilisation = if ($sku.PrepaidUnits.Enabled -gt 0) {
        [math]::Round(($sku.ConsumedUnits / $sku.PrepaidUnits.Enabled) * 100, 1)
    } else { 0 }

    $fields = @{
        "AuditDate"             = $auditDate
        "SKUPartNumber"         = $sku.SkuPartNumber
        "FriendlyName"          = $skuFriendlyNames[$sku.SkuPartNumber] ?? $sku.SkuPartNumber
        "Purchased"             = $sku.PrepaidUnits.Enabled
        "Assigned"              = $sku.ConsumedUnits
        "Utilisation"           = $utilisation
        "EstimatedMonthlySpend" = $sku.ConsumedUnits * ($skuMonthlyCost[$sku.SkuPartNumber] ?? 0)
    }
    New-MgSiteListItem -SiteId $siteId -ListId $summaryListId -Fields $fields
}

Write-Output "Inventory summary written to SharePoint."

# --- Step 2: Per-User Analysis ---
Write-Output "Retrieving all users..."
$users = Get-MgUser -All -Property @(
    "Id", "DisplayName", "UserPrincipalName", "AccountEnabled",
    "UserType", "AssignedLicenses", "SignInActivity", "Department"
)

Write-Output "Retrieved $($users.Count) users. Pulling usage report..."

$usagePath = Join-Path $env:TEMP "M365Usage_$(Get-Date -Format yyyyMMdd).csv"
$uri = "https://graph.microsoft.com/v1.0/reports/getOffice365ActiveUserDetail(period='D180')"
Invoke-MgGraphRequest -Uri $uri -OutputFilePath $usagePath
$usage = Import-Csv $usagePath

# --- Copilot usage (best-effort, beta endpoint) ---
$copilotActiveUPNs = @()
try {
    $copilotPath = Join-Path $env:TEMP "CopilotUsage_$(Get-Date -Format yyyyMMdd).csv"
    $copilotUri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='D180')"
    Invoke-MgGraphRequest -Uri $copilotUri -OutputFilePath $copilotPath
    $copilotUsage = Import-Csv $copilotPath
    $copilotActiveUPNs = ($copilotUsage | Where-Object { $_.'Last Activity Date' -ne "" }).'User Principal Name'
} catch {
    Write-Output "Copilot usage report not available: $_"
}

# --- Analyse each user ---
Write-Output "Analysing licence waste patterns..."
$wasteCount = 0

foreach ($user in $users) {
    foreach ($licence in $user.AssignedLicenses) {
        $skuName = ($skus | Where-Object { $_.SkuId -eq $licence.SkuId }).SkuPartNumber
        $monthlyCost = $skuMonthlyCost[$skuName] ?? 0

        # Skip free/zero-cost SKUs
        if ($monthlyCost -eq 0) { continue }

        # Determine waste category
        $wasteCategory = $null

        if (-not $user.AccountEnabled) {
            $wasteCategory = "Disabled Account"
        }
        elseif ($user.UserType -eq "Guest") {
            $wasteCategory = "Guest"
        }
        elseif ($user.UserPrincipalName -match $serviceAccountPattern) {
            $wasteCategory = "Service Account"
        }
        elseif ($skuName -match "Copilot" -and $user.UserPrincipalName -notin $copilotActiveUPNs) {
            $wasteCategory = "Copilot Unused"
        }
        else {
            # Check workload activity
            $activity = $usage | Where-Object {
                $_.'User Principal Name' -eq $user.UserPrincipalName
            }
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
                $wasteCategory = "Inactive User"
            }
        }

        # Write waste records to SharePoint
        if ($wasteCategory) {
            $fields = @{
                "AuditDate"            = $auditDate
                "UserPrincipalName"    = $user.UserPrincipalName
                "DisplayName"          = $user.DisplayName
                "Department"           = $user.Department ?? "Unset"
                "AccountEnabled"       = $user.AccountEnabled
                "UserType"             = $user.UserType
                "AssignedSKU"          = $skuName
                "SKUFriendlyName"      = $skuFriendlyNames[$skuName] ?? $skuName
                "WasteCategory"        = $wasteCategory
                "LastSignIn"           = $user.SignInActivity.LastSignInDateTime
                "LastWorkloadActivity" = $lastActivity
                "EstimatedMonthlyCost" = $monthlyCost
            }
            New-MgSiteListItem -SiteId $siteId -ListId $auditListId -Fields $fields
            $wasteCount++
        }
    }
}

Write-Output "Licence audit complete. $wasteCount waste records written to SharePoint."
