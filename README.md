# M365 Licence Audit

PowerShell scripts for auditing Microsoft 365 licence waste using Microsoft Graph API. Companion repository for [The M365 Licensing Audit Nobody Wants to Do](https://sbd.org.uk/blog/m365-licensing-audit).

## What This Does

Connects to your M365 tenant via Graph API and identifies:

- **Disabled accounts** still holding paid licences
- **Service accounts** on premium licences they don't need
- **Guest users** with paid licence assignments
- **Inactive users** with no workload activity in 90+ days
- **Copilot licences** assigned but never used
- **E5 users** who could be downgraded to E3

Outputs a waste report with estimated annual cost, and optionally writes results to SharePoint for a self-refreshing Power BI dashboard.

## Prerequisites

- PowerShell 7+
- [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation)
- Entra ID app registration with these **application** permissions (admin consent required):
  - `Organization.Read.All`
  - `User.Read.All`
  - `Reports.Read.All`
  - `Directory.Read.All`
  - `Sites.Manage.All` (only if writing to SharePoint)

> **Report privacy:** A Global Admin must disable user-level report obfuscation in **M365 Admin Centre > Settings > Org Settings > Reports** for the usage cross-reference to work. Without this, usage reports return hashed identifiers.

## Quick Start

```powershell
# Clone the repo
git clone https://github.com/wypbeu/m365-license-audit.git
cd m365-license-audit

# Connect to Graph (interactive — for one-off audits)
Connect-MgGraph -Scopes "Organization.Read.All","User.Read.All","Reports.Read.All","Directory.Read.All"

# Run the scripts in order
./scripts/Get-LicenceInventory.ps1 -OutputPath ./output
./scripts/Get-UserLicenceMap.ps1 -OutputPath ./output
./scripts/Find-LicenceWaste.ps1 -OutputPath ./output
./scripts/New-WasteReport.ps1 -OutputPath ./output
```

Results are written to `./output/` as CSV files.

## Repository Structure

```
├── scripts/
│   ├── Get-LicenceInventory.ps1     # Step 1: Pull subscribed SKUs
│   ├── Get-UserLicenceMap.ps1       # Step 2: Map per-user licence assignments
│   ├── Find-LicenceWaste.ps1        # Step 3: Identify waste patterns
│   └── New-WasteReport.ps1          # Step 4: Produce the consolidated waste report
├── runbook/
│   └── Invoke-LicenceAudit.ps1      # Azure Automation runbook (all-in-one)
├── config/
│   ├── sku-names.json               # SKU part number → friendly name mapping
│   └── sku-pricing.json             # SKU part number → monthly cost (GBP)
├── sharepoint/
│   └── list-schema.json             # SharePoint list column definitions
└── powerbi/
    └── measures.dax                 # DAX measures for the Power BI dashboard
```

## Scripts vs. Runbook

The **scripts/** folder contains individual scripts designed for interactive use — run them one at a time, inspect the output, iterate. They're what you use for a first-pass audit.

The **runbook/** contains a single consolidated script designed for Azure Automation. It runs unattended on a schedule, writes results to SharePoint, and feeds the Power BI dashboard. Deploy this once the interactive audit has proven its value.

## Configuration

Edit the files in `config/` to match your tenant:

- **sku-names.json** — Maps SKU part numbers to human-readable names. Microsoft's official mapping is incomplete; add any tenant-specific SKUs you encounter.
- **sku-pricing.json** — Monthly per-user cost in GBP for each SKU. Update these to match your Enterprise Agreement or CSP pricing.

## Azure Automation Deployment

See the [blog post](https://sbd.org.uk/blog/m365-licensing-audit) for full setup instructions. In summary:

1. Create an Azure Automation Account with a System-Assigned Managed Identity
2. Grant Graph permissions to the managed identity via `New-MgServicePrincipalAppRoleAssignment`
3. Create the SharePoint lists using the schema in `sharepoint/list-schema.json`
4. Import `runbook/Invoke-LicenceAudit.ps1` as a runbook
5. Update the configuration variables at the top of the runbook (site ID, list GUIDs)
6. Schedule to run weekly

## Licence

MIT
