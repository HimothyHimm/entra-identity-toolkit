<#
.SYNOPSIS
    Single authentication entry point for the entra-identity-toolkit.

.DESCRIPTION
    Connects to Microsoft Graph one of two ways:
      - APP-ONLY (unattended): when both -ClientId and -CertificateThumbprint are
        supplied, signs in as the toolkit's service principal using the certificate.
      - INTERACTIVE (delegated): otherwise, prompts for a user sign-in with the
        scopes the toolkit needs.

    Dot-source this from any toolkit script so auth lives in one place:

        . "$PSScriptRoot\Connect-ToolkitGraph.ps1" -ClientId $ClientId -CertificateThumbprint $Thumbprint

    Provision the app/cert these parameters refer to with New-ToolkitAppRegistration.ps1.

.PARAMETER TenantId
    Target tenant. Defaults to the demo tenant.

.PARAMETER ClientId
    App (client) ID of the toolkit's app registration. Triggers app-only auth when
    paired with -CertificateThumbprint.

.PARAMETER CertificateThumbprint
    Thumbprint of the signing certificate in CurrentUser\My.

.PARAMETER Scopes
    Delegated scopes used only for the interactive fallback.
#>
[CmdletBinding()]
param(
    [string]   $TenantId              = "ec5f5592-1d61-4f21-a7f9-3a49c1f78a58",
    [string]   $ClientId,
    [string]   $CertificateThumbprint,
    [string[]] $Scopes = @(
        "User.ReadWrite.All"
        "Group.ReadWrite.All"
        "Organization.Read.All"
        "Directory.Read.All"
        "RoleManagement.ReadWrite.Directory"
    )
)

if ($ClientId -and $CertificateThumbprint) {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint | Out-Null
    $ctx = Get-MgContext
    Write-Host "Connected APP-ONLY  ClientId=$($ctx.ClientId)  AuthType=$($ctx.AuthType)" -ForegroundColor Green
} else {
    Connect-MgGraph -TenantId $TenantId -Scopes $Scopes | Out-Null
    $ctx = Get-MgContext
    Write-Host "Connected INTERACTIVE  Account=$($ctx.Account)" -ForegroundColor Green
}
