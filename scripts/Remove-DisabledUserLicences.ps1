<#
.SYNOPSIS
    Reclaims licences from disabled accounts.

.DESCRIPTION
    Finds accounts that are disabled but still hold licence assignments and
    removes those licences. Supports -WhatIf and -Confirm; run with -WhatIf
    first to see what would be removed before removing anything.

    Requires an active Microsoft Graph connection with User.Read.All and
    LicenseAssignment.ReadWrite.All, the least privileged permission for
    licence writes.

.PARAMETER OutputPath
    Directory to write the removal log CSV. Defaults to ./output.

.EXAMPLE
    Connect-MgGraph -Scopes "User.Read.All","LicenseAssignment.ReadWrite.All"
    ./Remove-DisabledUserLicences.ps1 -WhatIf

.EXAMPLE
    ./Remove-DisabledUserLicences.ps1
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$OutputPath = (Join-Path $PSScriptRoot ".." "output")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Remove-DisabledAccountLicence {
    <#
    .SYNOPSIS
        Removes the named licences from a single user, honouring -WhatIf.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string[]]$LicencesToRemove
    )

    if ($PSCmdlet.ShouldProcess($UserPrincipalName, "Remove $($LicencesToRemove.Count) licence(s)")) {
        Set-MgUserLicense -UserId $UserId -AddLicenses @() -RemoveLicenses $LicencesToRemove
        Write-Host "Removed $($LicencesToRemove.Count) licence(s) from $UserPrincipalName" -ForegroundColor Green
        return $true
    }
    return $false
}

# Guard the main run so this script can be dot-sourced for testing without
# calling Get-MgUser or Set-MgUserLicense.
if ($MyInvocation.InvocationName -ne '.') {
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    Write-Host "Retrieving disabled accounts with licence assignments..." -ForegroundColor Cyan

    $disabled = Get-MgUser -All -Filter "accountEnabled eq false" `
        -Property "Id,UserPrincipalName,AssignedLicenses" |
        Where-Object { $_.AssignedLicenses.Count -gt 0 }

    $removalLog = [System.Collections.Generic.List[PSCustomObject]]::new()

    foreach ($user in $disabled) {
        $licencesToRemove = $user.AssignedLicenses | Select-Object -ExpandProperty SkuId

        $removed = Remove-DisabledAccountLicence -UserId $user.Id `
            -UserPrincipalName $user.UserPrincipalName `
            -LicencesToRemove $licencesToRemove

        # Under -WhatIf nothing is removed, so the log records candidates only.
        $removalLog.Add([PSCustomObject]@{
            UPN          = $user.UserPrincipalName
            Licences     = $licencesToRemove.Count
            Removed      = $removed
            CheckedAtUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        })
    }

    $logFile = Join-Path $OutputPath "licence-removal-log.csv"
    $removalLog | Export-Csv -Path $logFile -NoTypeInformation
    Write-Host "`nRemoval log written to $logFile" -ForegroundColor Green
}
