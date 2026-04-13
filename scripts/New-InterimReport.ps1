<#
.SYNOPSIS
    Produces an interim licence HTML report from inventory + user-licence map.

.DESCRIPTION
    Interim report for renewal conversations while waiting for the full waste
    analysis. Covers four of the five waste patterns using only data already
    on disk — no Graph connection required. Copilot adoption and workload
    activity require Find-LicenceWaste.ps1 and are flagged as pending.

    Inputs:
      - docs/licence-inventory.csv
      - docs/user-licence-map.csv
      - config/sku-pricing.json

    Output:
      - docs/interim-licence-report.html  (open in browser, print to PDF)

.EXAMPLE
    ./New-InterimReport.ps1
#>

[CmdletBinding()]
param(
    [string]$DocsPath   = (Join-Path $PSScriptRoot ".." "docs"),
    [string]$ConfigPath = (Join-Path $PSScriptRoot ".." "config"),
    [int]$InactiveDays  = 90,
    [string]$ServiceAccountPattern = "^(svc|service|admin|noreply|do-not-reply|mailbox|room|shared)"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$inventoryFile = Join-Path $DocsPath "licence-inventory.csv"
$mapFile       = Join-Path $DocsPath "user-licence-map.csv"
$pricingFile   = Join-Path $ConfigPath "sku-pricing.json"
$outputFile    = Join-Path $DocsPath "interim-licence-report.html"

foreach ($f in @($inventoryFile, $mapFile)) {
    if (-not (Test-Path $f)) { Write-Error "Missing input: $f"; return }
}

Write-Host "Loading inputs..." -ForegroundColor Cyan
$inventory = Import-Csv $inventoryFile
$users     = Import-Csv $mapFile

$pricing = @{}
if (Test-Path $pricingFile) {
    $pricing = Get-Content $pricingFile -Raw | ConvertFrom-Json -AsHashtable
    @($pricing.Keys) | Where-Object { $_ -like "_*" } | ForEach-Object { [void]$pricing.Remove($_) }
}

function Get-SkuMonthlyCost([string]$sku) {
    if ($pricing.ContainsKey($sku)) { return [double]$pricing[$sku] }
    return 0.0
}

function Format-Money([double]$n) { "£{0:N2}" -f $n }

$cutoff = (Get-Date).AddDays(-$InactiveDays)

# Parse LastSignIn (UK format DD/MM/YYYY HH:mm), blank = never
function Get-LastSignIn([string]$raw) {
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    $formats = @("dd/MM/yyyy HH:mm", "dd/MM/yyyy HH:mm:ss", "dd/MM/yyyy", "yyyy-MM-ddTHH:mm:ssZ", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd")
    foreach ($fmt in $formats) {
        try { return [datetime]::ParseExact($raw.Trim(), $fmt, [System.Globalization.CultureInfo]::InvariantCulture) } catch {}
    }
    try { return [datetime]$raw } catch { return $null }
}

# --- Inventory totals ---
Write-Host "Building inventory summary..." -ForegroundColor Cyan
$inventoryRows = $inventory | ForEach-Object {
    $sku  = $_.SKUPartNumber
    $unit = Get-SkuMonthlyCost $sku
    $assigned  = [int]$_.Assigned
    # Prefer pricing-driven calc; fall back to the CSV column if no pricing.
    $spend = if ($unit -gt 0) { $assigned * $unit } else { [double]$_.EstimatedMonthlySpend }
    [pscustomobject]@{
        SKU          = $sku
        Friendly     = $_.FriendlyName
        Purchased    = [int]$_.Purchased
        Assigned     = $assigned
        Available    = [int]$_.Available
        Utilisation  = [double]$_.Utilisation
        UnitCost     = $unit
        MonthlySpend = $spend
    }
}

# Exclude free trial / viral SKUs (no price) from headline totals — they skew idle counts.
$pricedRows     = @($inventoryRows | Where-Object { $_.UnitCost -gt 0 })
$totalPurchased = ($pricedRows | Measure-Object Purchased -Sum).Sum
$totalAssigned  = ($pricedRows | Measure-Object Assigned -Sum).Sum
$totalAvailable = ($pricedRows | Measure-Object Available -Sum).Sum
$totalSpend     = ($pricedRows | Measure-Object MonthlySpend -Sum).Sum
$unpricedSkus   = @($inventoryRows | Where-Object { $_.UnitCost -eq 0 -and $_.Assigned -gt 0 }).Count

# Idle spend = unassigned seats * pricing (only SKUs we have pricing for)
$idleSpend = 0.0
foreach ($r in $inventoryRows) {
    $unit = Get-SkuMonthlyCost $r.SKU
    $idleSpend += $r.Available * $unit
}

# --- Waste patterns ---
Write-Host "Analysing waste patterns..." -ForegroundColor Cyan

function Group-BySku($rows) {
    @($rows) |
        Group-Object SKU |
        ForEach-Object {
            $unit = Get-SkuMonthlyCost $_.Name
            [pscustomobject]@{
                SKU          = $_.Name
                Assignments  = $_.Count
                UnitCost     = $unit
                MonthlySpend = $_.Count * $unit
            }
        } |
        Sort-Object MonthlySpend -Descending
}

$disabledRows = @($users | Where-Object { $_.Enabled -eq "FALSE" })
$guestRows    = @($users | Where-Object { $_.UserType -eq "Guest" })
$svcRows      = @($users | Where-Object {
    $local = ($_.UPN -split "@")[0]
    $local -match $ServiceAccountPattern
})

$staleRows = @($users | Where-Object {
    if ($_.Enabled -ne "TRUE") { return $false }  # counted under disabled
    $dt = Get-LastSignIn $_.LastSignIn
    if ($null -eq $dt) { return $true }            # never signed in
    return $dt -lt $cutoff
})

$disabledBySku = Group-BySku $disabledRows
$guestBySku    = Group-BySku $guestRows
$svcBySku      = Group-BySku $svcRows
$staleBySku    = Group-BySku $staleRows

function Sum-Spend($rows) {
    if ($null -eq $rows -or @($rows).Count -eq 0) { return 0.0 }
    $m = @($rows) | Measure-Object MonthlySpend -Sum
    if ($null -eq $m.Sum) { return 0.0 } else { return [double]$m.Sum }
}

$disabledSpend = Sum-Spend $disabledBySku
$guestSpend    = Sum-Spend $guestBySku
$svcSpend      = Sum-Spend $svcBySku
$staleSpend    = Sum-Spend $staleBySku

$totalWasteMonthly = $disabledSpend + $guestSpend + $svcSpend + $staleSpend + $idleSpend
$totalWasteAnnual  = $totalWasteMonthly * 12

# --- HTML build ---
Write-Host "Rendering HTML..." -ForegroundColor Cyan

function Html-Encode([string]$s) {
    if ($null -eq $s) { return "" }
    [System.Net.WebUtility]::HtmlEncode($s)
}

function Build-InventoryTable($rows) {
    $sorted = $rows | Sort-Object MonthlySpend -Descending
    $body = foreach ($r in $sorted) {
        $utilClass = if ($r.Utilisation -lt 70) { "util-low" }
                     elseif ($r.Utilisation -lt 90) { "util-mid" }
                     else { "util-ok" }
        "<tr><td>$(Html-Encode $r.SKU)</td><td class='num'>$($r.Purchased)</td><td class='num'>$($r.Assigned)</td><td class='num'>$($r.Available)</td><td class='num $utilClass'>$("{0:N1}" -f $r.Utilisation)%</td><td class='num'>$(Format-Money $r.MonthlySpend)</td></tr>"
    }
    @"
<table>
  <thead><tr><th>SKU</th><th class='num'>Purchased</th><th class='num'>Assigned</th><th class='num'>Idle</th><th class='num'>Utilisation</th><th class='num'>Monthly spend</th></tr></thead>
  <tbody>
$($body -join "`n")
  </tbody>
</table>
"@
}

function Build-WasteTable($rows, $emptyMsg) {
    if ($null -eq $rows -or @($rows).Count -eq 0) { return "<p class='muted'>$emptyMsg</p>" }
    $body = foreach ($r in $rows) {
        "<tr><td>$(Html-Encode $r.SKU)</td><td class='num'>$($r.Assignments)</td><td class='num'>$(Format-Money $r.UnitCost)</td><td class='num'>$(Format-Money $r.MonthlySpend)</td></tr>"
    }
    @"
<table>
  <thead><tr><th>SKU</th><th class='num'>Assignments</th><th class='num'>Unit / mo</th><th class='num'>Monthly impact</th></tr></thead>
  <tbody>
$($body -join "`n")
  </tbody>
</table>
"@
}

$generated = Get-Date -Format "dd MMMM yyyy HH:mm"
$userCount = $users.Count
$uniqueUsers = ($users | Select-Object -Unique UPN | Measure-Object).Count

$html = @"
<!doctype html>
<html lang="en-GB">
<head>
<meta charset="utf-8">
<title>M365 Licence Audit — Interim Report</title>
<style>
  :root { --ink: #1a1a1a; --muted: #666; --line: #d8d8d8; --accent: #0b5394; --warn: #b45309; --bad: #b91c1c; --ok: #15803d; }
  * { box-sizing: border-box; }
  body { font-family: -apple-system, "Segoe UI", Helvetica, Arial, sans-serif; color: var(--ink); max-width: 1000px; margin: 2rem auto; padding: 0 1.5rem; line-height: 1.45; font-size: 13px; }
  h1 { font-size: 22px; margin: 0 0 .25rem; color: var(--accent); }
  h2 { font-size: 16px; margin: 2rem 0 .5rem; padding-bottom: .25rem; border-bottom: 2px solid var(--accent); }
  h3 { font-size: 14px; margin: 1.25rem 0 .5rem; color: var(--ink); }
  p.subtitle { color: var(--muted); margin: 0 0 1.5rem; }
  p.muted { color: var(--muted); font-style: italic; }
  .banner { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin: 1rem 0 1.5rem; }
  .stat { border: 1px solid var(--line); border-radius: 6px; padding: .75rem 1rem; background: #fafafa; }
  .stat .label { font-size: 11px; text-transform: uppercase; color: var(--muted); letter-spacing: .03em; }
  .stat .value { font-size: 18px; font-weight: 600; margin-top: .15rem; }
  .stat.bad .value  { color: var(--bad); }
  .stat.warn .value { color: var(--warn); }
  table { width: 100%; border-collapse: collapse; margin: .5rem 0 1rem; font-size: 12px; }
  th, td { padding: .4rem .6rem; border-bottom: 1px solid var(--line); text-align: left; }
  th { background: #f2f5f9; font-weight: 600; }
  td.num, th.num { text-align: right; font-variant-numeric: tabular-nums; }
  .util-low { color: var(--bad); font-weight: 600; }
  .util-mid { color: var(--warn); }
  .util-ok  { color: var(--ok); }
  .caveat { border-left: 3px solid var(--warn); background: #fff8eb; padding: .6rem 1rem; margin: 1rem 0; font-size: 12px; }
  footer { margin-top: 2rem; padding-top: 1rem; border-top: 1px solid var(--line); color: var(--muted); font-size: 11px; }
  @media print {
    body { margin: 0; padding: 0; font-size: 11px; max-width: none; }
    h2 { page-break-after: avoid; }
    table { page-break-inside: auto; }
    tr { page-break-inside: avoid; }
    .banner { grid-template-columns: repeat(4, 1fr); }
  }
</style>
</head>
<body>

<h1>M365 Licence Audit — Interim Report</h1>
<p class="subtitle">Generated $generated &middot; $uniqueUsers users &middot; $($inventoryRows.Count) SKUs</p>

<div class="caveat">
  <strong>Interim report.</strong> Copilot adoption and workload-level activity data are pending (require Graph usage reports).
  The inactive-user figure here is based on last sign-in only and may over- or under-state true inactivity.
  Full waste analysis will follow once the usage reports are available.
</div>

<div class="banner">
  <div class="stat"><div class="label">Total monthly spend</div><div class="value">$(Format-Money $totalSpend)</div></div>
  <div class="stat bad"><div class="label">Identified waste (mo)</div><div class="value">$(Format-Money $totalWasteMonthly)</div></div>
  <div class="stat bad"><div class="label">Identified waste (yr)</div><div class="value">$(Format-Money $totalWasteAnnual)</div></div>
  <div class="stat warn"><div class="label">Idle seats</div><div class="value">$("{0:N0}" -f $totalAvailable)</div></div>
</div>

<h2>1. Inventory &amp; utilisation (paid SKUs only)</h2>
<p>$totalPurchased seats purchased across $($pricedRows.Count) paid SKUs, $totalAssigned assigned, $totalAvailable idle. Unassigned paid seats represent <strong>$(Format-Money $idleSpend)/month</strong>. Microsoft's free trial / viral grants (e.g. <code>FORMS_PRO</code>, <code>FLOW_FREE</code>) are excluded from this table — they show as "purchased" on the tenant but carry no cost. $unpricedSkus assigned SKU(s) have no entry in <code>sku-pricing.json</code> and are also excluded; treat the headline figure as a floor estimate.</p>
$(Build-InventoryTable $pricedRows)

<h2>2. Waste patterns</h2>

<h3>2.1 Disabled accounts holding licences</h3>
<p>$($disabledRows.Count) assignments on disabled accounts — <strong>$(Format-Money $disabledSpend)/month</strong>. These should be removed before renewal.</p>
$(Build-WasteTable $disabledBySku "No disabled accounts hold licences.")

<h3>2.2 Guest users with paid licences</h3>
<p>$($guestRows.Count) assignments on guest accounts — <strong>$(Format-Money $guestSpend)/month</strong>. Guests rarely need paid SKUs.</p>
$(Build-WasteTable $guestBySku "No guest users hold paid licences.")

<h3>2.3 Likely service accounts on paid SKUs</h3>
<p>$($svcRows.Count) assignments on accounts matching the service-account pattern — <strong>$(Format-Money $svcSpend)/month</strong>. Review individually; some may need shared mailbox conversion.</p>
$(Build-WasteTable $svcBySku "No service accounts identified.")

<h3>2.4 Inactive users (no sign-in in $InactiveDays+ days)</h3>
<p>$($staleRows.Count) enabled accounts with no recent sign-in — <strong>$(Format-Money $staleSpend)/month</strong>. Confirm with HR/joiners-leavers before reclaiming.</p>
$(Build-WasteTable $staleBySku "No inactive users identified.")

<h3>2.5 Copilot adoption</h3>
<p class="muted">Pending — requires Graph usage report (Reports.Read.All).</p>

<h2>3. Renewal recommendations</h2>
<ul>
  <li>Reclaim the $($disabledRows.Count) licences on disabled accounts before the renewal baseline is set.</li>
  <li>Review the $($staleRows.Count) inactive enabled users — even a 50% true-positive rate is material at this scale.</li>
  <li>Reduce unassigned seats on any SKU where utilisation is below 70% and the gap is not earmarked for near-term growth.</li>
  <li>Re-run this report once the full waste analysis (Copilot + workload activity) is in hand to firm up the annual saving.</li>
</ul>

<footer>
  Generated by <code>New-InterimReport.ps1</code> from <code>licence-inventory.csv</code> and <code>user-licence-map.csv</code>.
  Pricing sourced from <code>config/sku-pricing.json</code> — validate against current EA/CSP rates before quoting externally.
</footer>

</body>
</html>
"@

$html | Set-Content -Path $outputFile -Encoding UTF8
Write-Host "Report written to: $outputFile" -ForegroundColor Green
Write-Host ""
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host ("  Total monthly spend:   {0}" -f (Format-Money $totalSpend))
Write-Host ("  Identified waste (mo): {0}" -f (Format-Money $totalWasteMonthly))
Write-Host ("  Identified waste (yr): {0}" -f (Format-Money $totalWasteAnnual))
Write-Host ("  Disabled assignments:  {0} ({1}/mo)" -f $disabledRows.Count, (Format-Money $disabledSpend))
Write-Host ("  Guest assignments:     {0} ({1}/mo)" -f $guestRows.Count,    (Format-Money $guestSpend))
Write-Host ("  Service-account rows:  {0} ({1}/mo)" -f $svcRows.Count,      (Format-Money $svcSpend))
Write-Host ("  Inactive enabled:      {0} ({1}/mo)" -f $staleRows.Count,    (Format-Money $staleSpend))
