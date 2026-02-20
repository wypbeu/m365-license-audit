<#
.SYNOPSIS
    Produces a consolidated licence waste report from the waste analysis data.

.DESCRIPTION
    Reads the waste CSV from Find-LicenceWaste.ps1 and produces a summary report
    grouped by waste category, with estimated monthly and annual costs.
    Outputs both to the console and as CSV files.

.PARAMETER OutputPath
    Directory containing licence-waste.csv and where the report is written.
    Defaults to ./output.

.EXAMPLE
    ./New-WasteReport.ps1 -OutputPath ./output
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = (Join-Path $PSScriptRoot ".." "output")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Load waste data ---
$wasteFile = Join-Path $OutputPath "licence-waste.csv"
if (-not (Test-Path $wasteFile)) {
    Write-Error "Waste data not found at $wasteFile — run Find-LicenceWaste.ps1 first"
    return
}
$waste = Import-Csv $wasteFile

if ($waste.Count -eq 0) {
    Write-Host "No licence waste found. Your tenant is in good shape." -ForegroundColor Green
    return
}

# --- Summary by category ---
Write-Host "`n=== LICENCE WASTE REPORT ===" -ForegroundColor Cyan
Write-Host "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n"

$categorySummary = $waste |
    Group-Object WasteCategory |
    Select-Object @{
        Name = "Category"; Expression = { $_.Name }
    }, @{
        Name = "Users"; Expression = { ($_.Group | Select-Object -Unique UPN | Measure-Object).Count }
    }, @{
        Name = "Assignments"; Expression = { $_.Count }
    }, @{
        Name = "MonthlyCost"; Expression = {
            ($_.Group | ForEach-Object { [decimal]$_.MonthlyCost } | Measure-Object -Sum).Sum
        }
    }, @{
        Name = "AnnualCost"; Expression = {
            ($_.Group | ForEach-Object { [decimal]$_.MonthlyCost } | Measure-Object -Sum).Sum * 12
        }
    } |
    Sort-Object AnnualCost -Descending

$categorySummary | Format-Table Category, Users, Assignments, @{
    Name = "Monthly (GBP)"; Expression = { "£{0:N0}" -f $_.MonthlyCost }; Align = "Right"
}, @{
    Name = "Annual (GBP)"; Expression = { "£{0:N0}" -f $_.AnnualCost }; Align = "Right"
} -AutoSize

$totalMonthly = ($categorySummary | Measure-Object -Property MonthlyCost -Sum).Sum
$totalAnnual = $totalMonthly * 12

Write-Host "Total estimated monthly waste: £$('{0:N0}' -f $totalMonthly)" -ForegroundColor Red
Write-Host "Total estimated annual waste:  £$('{0:N0}' -f $totalAnnual)" -ForegroundColor Red

# --- Summary by SKU ---
Write-Host "`n--- Waste by Licence Type ---" -ForegroundColor Yellow

$skuSummary = $waste |
    Group-Object FriendlyName |
    Select-Object @{
        Name = "Licence"; Expression = { $_.Name }
    }, @{
        Name = "WastedAssignments"; Expression = { $_.Count }
    }, @{
        Name = "MonthlyCost"; Expression = {
            ($_.Group | ForEach-Object { [decimal]$_.MonthlyCost } | Measure-Object -Sum).Sum
        }
    } |
    Sort-Object MonthlyCost -Descending

$skuSummary | Format-Table Licence, WastedAssignments, @{
    Name = "Monthly (GBP)"; Expression = { "£{0:N0}" -f $_.MonthlyCost }; Align = "Right"
} -AutoSize

# --- Top departments ---
Write-Host "--- Top Departments by Waste ---" -ForegroundColor Yellow

$deptSummary = $waste |
    Where-Object { $_.Department -ne "" } |
    Group-Object Department |
    Select-Object @{
        Name = "Department"; Expression = { $_.Name }
    }, @{
        Name = "WastedAssignments"; Expression = { $_.Count }
    }, @{
        Name = "MonthlyCost"; Expression = {
            ($_.Group | ForEach-Object { [decimal]$_.MonthlyCost } | Measure-Object -Sum).Sum
        }
    } |
    Sort-Object MonthlyCost -Descending |
    Select-Object -First 10

$deptSummary | Format-Table Department, WastedAssignments, @{
    Name = "Monthly (GBP)"; Expression = { "£{0:N0}" -f $_.MonthlyCost }; Align = "Right"
} -AutoSize

# --- Export reports ---
$categoryFile = Join-Path $OutputPath "waste-report-by-category.csv"
$categorySummary | Export-Csv -Path $categoryFile -NoTypeInformation

$skuFile = Join-Path $OutputPath "waste-report-by-sku.csv"
$skuSummary | Export-Csv -Path $skuFile -NoTypeInformation

$deptFile = Join-Path $OutputPath "waste-report-by-department.csv"
$deptSummary | Export-Csv -Path $deptFile -NoTypeInformation

Write-Host "`nReports written to:" -ForegroundColor Green
Write-Host "  $categoryFile"
Write-Host "  $skuFile"
Write-Host "  $deptFile"
