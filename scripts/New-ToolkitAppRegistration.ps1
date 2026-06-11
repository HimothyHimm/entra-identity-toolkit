<#
.SYNOPSIS
    Provisions the entra-identity-toolkit's own Microsoft Entra app registration:
    application + service principal, a self-signed signing certificate, and the
    Microsoft Graph application permissions it needs (admin-consented). After this
    runs once, every other toolkit script can authenticate as a service principal
    with the certificate instead of an interactive user.

.DESCRIPTION
    Idempotent and safe to re-run. Reuses an existing app, service principal, and
    certificate when present, and skips any permission already granted. All writes
    go through raw Microsoft Graph (Invoke-MgGraphRequest), consistent with the rest
    of the toolkit and the SDK quirks documented in the README.

    You run this ONCE, signed in as a Global Administrator (the bootstrap step needs
    a human with consent rights). Afterward the toolkit runs unattended off the cert.

.PARAMETER TenantId
    The Entra tenant to provision in. Defaults to the demo tenant.

.PARAMETER AppDisplayName
    Display name for the app registration.

.PARAMETER CertSubject
    Subject (CN) for the self-signed signing certificate, stored in CurrentUser\My.

.PARAMETER Permissions
    Microsoft Graph APPLICATION permissions to grant, by value.

.NOTES
    RoleManagement.ReadWrite.Directory is the highest-privilege grant in the default
    set. It is here because the toolkit performs directory role assignment (the
    cloud-admin provisioning capability). In a production tenant you would justify
    or narrow this deliberately rather than grant it by default.

.EXAMPLE
    .\New-ToolkitAppRegistration.ps1
    # Provisions everything with defaults and prints the ready-to-use connect line.
#>
[CmdletBinding()]
param(
    [string]   $TenantId       = "ec5f5592-1d61-4f21-a7f9-3a49c1f78a58",
    [string]   $AppDisplayName = "entra-identity-toolkit",
    [string]   $CertSubject    = "entra-identity-toolkit",
    [string[]] $Permissions    = @(
        "User.ReadWrite.All"
        "Group.ReadWrite.All"
        "Organization.Read.All"
        "Directory.Read.All"
        "RoleManagement.ReadWrite.Directory"
    )
)

$ErrorActionPreference = "Stop"
$GraphBase = "https://graph.microsoft.com/v1.0"

# --- 1. Connect as admin (delegated) with the scopes needed to provision an app ---
Connect-MgGraph -TenantId $TenantId -Scopes @(
    "Application.ReadWrite.All"
    "AppRoleAssignment.ReadWrite.All"
) | Out-Null
Write-Host "Connected as $((Get-MgContext).Account)" -ForegroundColor Cyan

# --- 2. App registration (create or reuse) ---
$app = (Invoke-MgGraphRequest -Method GET -Uri "$GraphBase/applications?`$filter=displayName eq '$AppDisplayName'").value |
       Select-Object -First 1
if ($app) {
    Write-Host "APP  reuse  $($app.appId)" -ForegroundColor Yellow
} else {
    $app = Invoke-MgGraphRequest -Method POST -Uri "$GraphBase/applications" `
        -Body (@{ displayName = $AppDisplayName } | ConvertTo-Json) -ContentType "application/json"
    Write-Host "APP  create $($app.appId)" -ForegroundColor Green
}

# --- 3. Service principal (create or reuse) ---
$sp = (Invoke-MgGraphRequest -Method GET -Uri "$GraphBase/servicePrincipals?`$filter=appId eq '$($app.appId)'").value |
      Select-Object -First 1
if ($sp) {
    Write-Host "SP   reuse  $($sp.id)" -ForegroundColor Yellow
} else {
    $sp = Invoke-MgGraphRequest -Method POST -Uri "$GraphBase/servicePrincipals" `
        -Body (@{ appId = $app.appId } | ConvertTo-Json) -ContentType "application/json"
    Write-Host "SP   create $($sp.id)" -ForegroundColor Green
}

# --- 4. Signing certificate (reuse newest valid, else create) ---
$cert = Get-ChildItem "Cert:\CurrentUser\My" |
        Where-Object { $_.Subject -eq "CN=$CertSubject" -and $_.NotAfter -gt (Get-Date) } |
        Sort-Object NotAfter -Descending | Select-Object -First 1
if ($cert) {
    Write-Host "CERT reuse  $($cert.Thumbprint)" -ForegroundColor Yellow
} else {
    $cert = New-SelfSignedCertificate -Subject "CN=$CertSubject" `
        -CertStoreLocation "Cert:\CurrentUser\My" `
        -KeyExportPolicy Exportable -KeySpec Signature -NotAfter (Get-Date).AddYears(1)
    Write-Host "CERT create $($cert.Thumbprint)" -ForegroundColor Green
}

# Graph stores customKeyIdentifier as base64 of the cert thumbprint bytes; build it to compare
$thumbBytes = for ($i = 0; $i -lt $cert.Thumbprint.Length; $i += 2) {
    [Convert]::ToByte($cert.Thumbprint.Substring($i, 2), 16)
}
$cki = [Convert]::ToBase64String([byte[]]$thumbBytes)

$appKeys = (Invoke-MgGraphRequest -Method GET -Uri "$GraphBase/applications/$($app.id)?`$select=keyCredentials").keyCredentials
if ($appKeys | Where-Object { $_.customKeyIdentifier -eq $cki }) {
    Write-Host "CERT already on app" -ForegroundColor Yellow
} else {
    $keyBody = @{
        keyCredentials = @(@{
            type        = "AsymmetricX509Cert"
            usage       = "Verify"
            key         = [Convert]::ToBase64String($cert.RawData)
            displayName = "CN=$CertSubject"
        })
    } | ConvertTo-Json -Depth 5
    Invoke-MgGraphRequest -Method PATCH -Uri "$GraphBase/applications/$($app.id)" `
        -Body $keyBody -ContentType "application/json" | Out-Null
    Write-Host "CERT uploaded to app" -ForegroundColor Green
}

# --- 5. Grant Graph application permissions (appRoleAssignments == admin consent) ---
$graphSp = (Invoke-MgGraphRequest -Method GET `
    -Uri "$GraphBase/servicePrincipals?`$filter=appId eq '00000003-0000-0000-c000-000000000000'").value |
    Select-Object -First 1
$existing = (Invoke-MgGraphRequest -Method GET -Uri "$GraphBase/servicePrincipals/$($sp.id)/appRoleAssignments").value

foreach ($permValue in $Permissions) {
    $role = $graphSp.appRoles |
            Where-Object { $_.value -eq $permValue -and $_.allowedMemberTypes -contains "Application" } |
            Select-Object -First 1
    if (-not $role) {
        Write-Host "MISS no application role found for $permValue" -ForegroundColor Red
        continue
    }
    if ($existing | Where-Object { $_.appRoleId -eq $role.id }) {
        Write-Host "PERM have   $permValue" -ForegroundColor Yellow
        continue
    }
    $assignBody = @{ principalId = $sp.id; resourceId = $graphSp.id; appRoleId = $role.id } | ConvertTo-Json
    Invoke-MgGraphRequest -Method POST `
        -Uri "$GraphBase/servicePrincipals/$($sp.id)/appRoleAssignments" `
        -Body $assignBody -ContentType "application/json" | Out-Null
    Write-Host "PERM grant  $permValue" -ForegroundColor Green
}

# --- 6. Summary + ready-to-use connection line ---
Write-Host ""
Write-Host "=== entra-identity-toolkit app registration ready ===" -ForegroundColor Cyan
Write-Host "TenantId   : $TenantId"
Write-Host "ClientId   : $($app.appId)"
Write-Host "Thumbprint : $($cert.Thumbprint)"
Write-Host ""
Write-Host "Authenticate the toolkit unattended with:" -ForegroundColor Cyan
Write-Host "  Connect-MgGraph -TenantId $TenantId -ClientId $($app.appId) -CertificateThumbprint $($cert.Thumbprint)"
Write-Host ""
Write-Host "Note: appRoleAssignments can take 1-2 minutes to propagate before app-only" -ForegroundColor DarkGray
Write-Host "      calls succeed. A 403 immediately after a fresh grant is expected; wait." -ForegroundColor DarkGray
