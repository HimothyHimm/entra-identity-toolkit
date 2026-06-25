<#
.SYNOPSIS
    Bulk-creates Microsoft Entra ID users from a CSV (demo-data seeding / onboarding).

.DESCRIPTION
    Reads a CSV of users and creates each one in Entra ID via Microsoft Graph.
      - Connects via Connect-ToolkitGraph.ps1 (app-only if ClientId+thumbprint
        supplied, interactive otherwise), pinned to the org tenant.
      - Resolves the tenant's default verified domain automatically (override with -Domain).
      - Builds userPrincipalName and mailNickname as first.last@<domain>.
      - Sets a random strong temporary password with force-change at next sign-in.
      - Is idempotent: users that already exist are skipped, not re-created.
      - A second pass wires up manager relationships using the raw Graph API.

    Note on the manager pass: the Set-MgUserManagerByRef cmdlet silently no-ops in some
    SDK builds (accepts the call, writes nothing). This script uses Invoke-MgGraphRequest
    to PUT the manager/$ref directly, which writes reliably.

.PARAMETER CsvPath
    Path to the input CSV. Columns:
    FirstName,LastName,JobTitle,Department,UsageLocation,Manager
    (Manager holds the manager's first.last key, or is blank for top-level staff.)

.PARAMETER TenantId
    Entra tenant to connect to. Defaults to this toolkit's tenant.

.PARAMETER Domain
    UPN domain. Defaults to the tenant's default verified domain.

.PARAMETER ClientId
    App (client) ID of the toolkit's app registration. Supply with
    -CertificateThumbprint to run unattended (app-only). Omit for interactive.

.PARAMETER CertificateThumbprint
    Thumbprint of the signing certificate in CurrentUser\My (app-only auth).

.EXAMPLE
    .\New-DemoUsers.ps1 -CsvPath ..\data\sample-users.csv

.EXAMPLE
    # Dry run -- shows what would be created without touching the tenant
    .\New-DemoUsers.ps1 -CsvPath ..\data\sample-users.csv -WhatIf
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$CsvPath,

    [string]$TenantId = "ec5f5592-1d61-4f21-a7f9-3a49c1f78a58",

    [string]$Domain,

    [string]$ClientId,

    [string]$CertificateThumbprint
)

$ErrorActionPreference = 'Stop'

# --- Connect (app-only if ClientId+thumbprint supplied, otherwise interactive) ---
. "$PSScriptRoot\Connect-ToolkitGraph.ps1" `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -CertificateThumbprint $CertificateThumbprint `
    -Scopes @('User.ReadWrite.All')

# --- Resolve the UPN domain from the tenant if not supplied ---
if (-not $Domain) {
    $verified = (Get-MgOrganization).VerifiedDomains
    $Domain = ($verified | Where-Object { $_.IsDefault }).Name
    if (-not $Domain) { $Domain = ($verified | Where-Object { $_.IsInitial }).Name }
}
Write-Host "Using domain: $Domain`n" -ForegroundColor Cyan

# --- Helper: strong random temporary password (meets Entra complexity) ---
function New-TempPassword {
    $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789!@#$%*?'
    -join (1..16 | ForEach-Object { $chars[(Get-Random -Maximum $chars.Length)] })
}

# --- Load CSV ---
if (-not (Test-Path $CsvPath)) { throw "CSV not found: $CsvPath" }
$rows = Import-Csv -Path $CsvPath
Write-Host "Loaded $($rows.Count) rows from $CsvPath`n" -ForegroundColor Cyan

$created = 0; $skipped = 0; $failed = 0
$keyToUpn = @{}   # first.last -> full UPN, used to resolve managers in pass 2

# --- Pass 1: create users ---
foreach ($row in $rows) {
    $key     = ("{0}.{1}" -f $row.FirstName, $row.LastName).ToLower().Replace(' ', '')
    $upn     = "$key@$Domain"
    $display = "$($row.FirstName) $($row.LastName)"
    $keyToUpn[$key] = $upn

    # Idempotency: skip if the user already exists
    $existing = Get-MgUser -Filter "userPrincipalName eq '$upn'" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "SKIP   $display ($upn) - already exists" -ForegroundColor DarkYellow
        $skipped++
        continue
    }

    $usage = if ($row.UsageLocation) { $row.UsageLocation } else { 'US' }

    $params = @{
        AccountEnabled    = $true
        DisplayName       = $display
        GivenName         = $row.FirstName
        Surname           = $row.LastName
        UserPrincipalName = $upn
        MailNickname      = $key
        JobTitle          = $row.JobTitle
        Department        = $row.Department
        UsageLocation     = $usage
        PasswordProfile   = @{
            Password                      = New-TempPassword
            ForceChangePasswordNextSignIn = $true
        }
    }

    try {
        if ($PSCmdlet.ShouldProcess($upn, 'Create user')) {
            New-MgUser @params | Out-Null
            Write-Host "CREATE $display ($upn)" -ForegroundColor Green
            $created++
        }
    }
    catch {
        Write-Host "FAIL   $display ($upn): $($_.Exception.Message)" -ForegroundColor Red
        $failed++
    }
}

# --- Pass 2: assign managers via raw Graph (Set-MgUserManagerByRef no-ops in some SDK builds) ---
Write-Host "`nAssigning managers..." -ForegroundColor Cyan
Start-Sleep -Seconds 5   # let just-created accounts replicate before we reference them

# Build a UPN -> Id map once, so we never depend on per-row filter consistency
$idByUpn = @{}
Get-MgUser -All | ForEach-Object { $idByUpn[$_.UserPrincipalName.ToLower()] = $_.Id }

foreach ($row in $rows) {
    if (-not $row.Manager) { continue }

    $key        = ("{0}.{1}" -f $row.FirstName, $row.LastName).ToLower().Replace(' ', '')
    $display    = "$($row.FirstName) $($row.LastName)"
    $userUpn    = ("$key@$Domain").ToLower()
    $managerUpn = ("$($row.Manager.Replace(' ', ''))@$Domain").ToLower()
    $userId     = $idByUpn[$userUpn]
    $managerId  = $idByUpn[$managerUpn]

    if (-not $userId -or -not $managerId) {
        Write-Host "WARN   missing id for $userUpn or $managerUpn" -ForegroundColor DarkYellow
        continue
    }

    if ($PSCmdlet.ShouldProcess($userUpn, "Set manager -> $managerUpn")) {
        $body = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/users/$managerId" } | ConvertTo-Json
        try {
            Invoke-MgGraphRequest -Method PUT `
                -Uri "https://graph.microsoft.com/v1.0/users/$userId/manager/`$ref" `
                -Body $body -ContentType "application/json" -ErrorAction Stop
            Write-Host "MANAGER $display -> $managerUpn" -ForegroundColor Green
        }
        catch {
            Write-Host "WARN   could not set manager for $userUpn -> $managerUpn : $($_.Exception.Message)" -ForegroundColor DarkYellow
        }
    }
}

# --- Summary ---
Write-Host "`nDone.  Created: $created   Skipped: $skipped   Failed: $failed" -ForegroundColor Cyan