<#
.SYNOPSIS
    Generates a directory report (org chart, account inventory, license status)
    from Microsoft Entra ID and exports it to CSV and HTML.

.DESCRIPTION
    Pulls all users via raw Graph with the manager expanded inline, plus account
    status, department, created date, and any assigned licenses. Rolls up
    department headcount and writes a timestamped CSV and a self-contained HTML
    report to .\reports.

    Connects via Connect-ToolkitGraph.ps1 (app-only if ClientId+thumbprint
    supplied, interactive otherwise) with read-only scopes.

    Uses raw Graph (Invoke-MgGraphRequest) throughout. The manager relationship
    is read via the $expand query that proved correct in Stage 1, because the
    SDK's *Manager cmdlets are unreliable in this environment.

    Last sign-in (signInActivity) is intentionally omitted: it requires Entra ID
    P1/P2. This report is designed to run on the free tier.

.PARAMETER OutputDir
    Folder for the CSV and HTML output. Defaults to .\reports.

.PARAMETER TenantId
    Entra tenant to connect to. Defaults to this toolkit's tenant.

.PARAMETER ClientId
    App (client) ID of the toolkit's app registration. Supply with
    -CertificateThumbprint to run unattended (app-only). Omit for interactive.

.PARAMETER CertificateThumbprint
    Thumbprint of the signing certificate in CurrentUser\My (app-only auth).

.EXAMPLE
    .\Get-DirectoryReport.ps1
    .\Get-DirectoryReport.ps1 -OutputDir .\reports
#>
[CmdletBinding()]
param(
    [string]$OutputDir = ".\reports",

    [string]$TenantId = "ec5f5592-1d61-4f21-a7f9-3a49c1f78a58",

    [string]$ClientId,

    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'

# --- Connect (app-only if ClientId+thumbprint supplied, otherwise interactive) ---
. "$PSScriptRoot\Connect-ToolkitGraph.ps1" `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint `
    -Scopes @('User.Read.All','Organization.Read.All')

# --- Pull users, manager expanded inline (raw Graph - the layer we trust) -----
$select = "displayName,userPrincipalName,department,jobTitle,accountEnabled,createdDateTime,assignedLicenses"
$uri = "https://graph.microsoft.com/v1.0/users?`$select=$select&`$expand=manager(`$select=displayName)&`$top=100"
$users = @()
do {
    $resp  = Invoke-MgGraphRequest -Method GET -Uri $uri
    $users += $resp.value
    $uri   = $resp.'@odata.nextLink'
} while ($uri)

# --- SkuId -> friendly name map (for any assigned licenses) -------------------
$skuName = @{}
try { Get-MgSubscribedSku | ForEach-Object { $skuName[$_.SkuId] = $_.SkuPartNumber } } catch {}

# --- Shape rows ---------------------------------------------------------------
$report = $users | ForEach-Object {
    $lic = if ($_.assignedLicenses -and $_.assignedLicenses.Count) {
        ($_.assignedLicenses | ForEach-Object {
            if ($skuName.ContainsKey($_.skuId)) { $skuName[$_.skuId] } else { $_.skuId }
        }) -join ', '
    } else { '(none)' }

    [PSCustomObject]@{
        DisplayName = $_.displayName
        UPN         = $_.userPrincipalName
        Department  = if ($_.department) { $_.department } else { '(none)' }
        JobTitle    = $_.jobTitle
        Manager     = if ($_.manager) { $_.manager.displayName } else { '(none)' }
        Enabled     = $_.accountEnabled
        Created     = if ($_.createdDateTime) { ([datetime]$_.createdDateTime).ToString('yyyy-MM-dd') } else { '' }
        Licenses    = $lic
    }
} | Sort-Object Department, DisplayName

# --- Department rollup --------------------------------------------------------
$deptRollup = $report | Where-Object { $_.Department -ne '(none)' } |
    Group-Object Department | Sort-Object Name |
    ForEach-Object { [PSCustomObject]@{ Department = $_.Name; Headcount = $_.Count } }

# --- Console summary ----------------------------------------------------------
Write-Host ("Users: {0}   Departments: {1}" -f $report.Count, $deptRollup.Count) -ForegroundColor Cyan
$report | Format-Table DisplayName, Department, Manager, Enabled, Licenses -AutoSize

# --- Export -------------------------------------------------------------------
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
$stamp    = Get-Date -Format 'yyyyMMdd-HHmmss'
$csvPath  = Join-Path $OutputDir "directory-$stamp.csv"
$htmlPath = Join-Path $OutputDir "directory-$stamp.html"

$report | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

$style = @"
<style>
 body { font-family: 'Segoe UI', Arial, sans-serif; margin: 2rem; color: #1b1b1f; }
 h1 { font-size: 1.4rem; } h2 { font-size: 1.05rem; margin-top: 2rem; }
 table { border-collapse: collapse; width: 100%; margin-top: .5rem; font-size: .9rem; }
 th, td { border: 1px solid #d0d0d7; padding: 6px 10px; text-align: left; }
 th { background: #f3f2f8; } tr:nth-child(even) td { background: #fafafd; }
 .meta { color: #5b5b66; font-size: .85rem; }
</style>
"@
$genAt      = Get-Date -Format 'yyyy-MM-dd HH:mm'
$rollupHtml = $deptRollup | ConvertTo-Html -Fragment
$tableHtml  = $report | ConvertTo-Html -Fragment
$html = @"
<!doctype html><html><head><meta charset="utf-8"><title>Directory Report</title>$style</head><body>
<h1>Entra ID Directory Report</h1>
<p class="meta">Generated $genAt &middot; $($report.Count) users &middot; $($deptRollup.Count) departments</p>
<h2>Department headcount</h2>
$rollupHtml
<h2>Users</h2>
$tableHtml
</body></html>
"@
$html | Out-File -FilePath $htmlPath -Encoding UTF8

Write-Host "CSV : $csvPath"  -ForegroundColor Green
Write-Host "HTML: $htmlPath" -ForegroundColor Green