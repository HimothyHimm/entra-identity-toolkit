<#
.SYNOPSIS
    Creates one Entra ID security group per department and adds each user to the
    group matching their Department attribute.

.DESCRIPTION
    Reads all users from the connected tenant, groups them by their Department
    attribute, ensures a security group exists for each department, and adds the
    matching users as members.

    Idempotent  - existing groups are reused; users already in a group are skipped.
    Dry run     - supports -WhatIf to preview without changing anything.

    Group creation, membership writes, and member read-backs all go through raw
    Graph (Invoke-MgGraphRequest). In this SDK build the *ByRef / reference
    cmdlets accept the call and silently do nothing, so raw Graph is the only
    reliable path for reference writes (this was proven during Stage 1's manager
    assignment).

.PARAMETER GroupPrefix
    Prefix for the group display name. Default: "Department - ".

.EXAMPLE
    .\New-DepartmentGroups.ps1 -WhatIf
    .\New-DepartmentGroups.ps1
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$GroupPrefix = "Department - "
)

$ErrorActionPreference = 'Stop'

function Get-ExistingGroup {
    param([string]$MailNickname)
    $uri  = "https://graph.microsoft.com/v1.0/groups?`$filter=mailNickname eq '$MailNickname'&`$select=id,displayName,mailNickname"
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
    if ($resp.value -and $resp.value.Count -gt 0) { return $resp.value[0] }
    return $null
}

function Get-GroupMemberIds {
    param([string]$GroupId)
    $ids = New-Object System.Collections.Generic.HashSet[string]
    $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=id"
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
        foreach ($m in $resp.value) { [void]$ids.Add($m.id) }
        $uri = $resp.'@odata.nextLink'
    } while ($uri)
    return ,$ids
}

# --- 1. Pull users that have a Department set --------------------------------
$users = Get-MgUser -All -Property Id,DisplayName,Department,UserPrincipalName |
         Where-Object { $_.Department }

if (-not $users) {
    Write-Host "No users with a Department attribute found. Run New-DemoUsers.ps1 first." -ForegroundColor Yellow
    return
}

$byDept = $users | Group-Object Department | Sort-Object Name
Write-Host ("Found {0} users across {1} departments." -f $users.Count, $byDept.Count) -ForegroundColor Cyan

# --- 2. Ensure a group per department, then add members ----------------------
foreach ($dept in $byDept) {
    $deptName     = $dept.Name
    $displayName  = "$GroupPrefix$deptName"
    $mailNickname = ("dept-" + ($deptName -replace '[^a-zA-Z0-9]', '')).ToLower()

    $group = Get-ExistingGroup -MailNickname $mailNickname

    if ($group) {
        Write-Host "GROUP  exists  $displayName" -ForegroundColor DarkGray
    }
    else {
        if ($PSCmdlet.ShouldProcess($displayName, "Create security group")) {
            $body = @{
                displayName     = $displayName
                description     = "Members of the $deptName department (managed by entra-identity-toolkit)."
                mailEnabled     = $false
                mailNickname    = $mailNickname
                securityEnabled = $true
            } | ConvertTo-Json
            $group = Invoke-MgGraphRequest -Method POST `
                        -Uri "https://graph.microsoft.com/v1.0/groups" `
                        -Body $body -ContentType "application/json"
            Write-Host "GROUP  created $displayName" -ForegroundColor Green
            Start-Sleep -Seconds 5   # let the new group replicate before adding members
        }
        else {
            Write-Host "What if: would create group '$displayName' and add $($dept.Count) members." -ForegroundColor Yellow
            continue
        }
    }

    $existing = Get-GroupMemberIds -GroupId $group.id

    foreach ($u in $dept.Group) {
        if ($existing.Contains($u.Id)) {
            Write-Host "  SKIP $($u.DisplayName) - already a member" -ForegroundColor DarkGray
            continue
        }
        if ($PSCmdlet.ShouldProcess($u.DisplayName, "Add to $displayName")) {
            $refBody = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($u.Id)" } | ConvertTo-Json
            Invoke-MgGraphRequest -Method POST `
                -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/members/`$ref" `
                -Body $refBody -ContentType "application/json"
            Write-Host "  ADD  $($u.DisplayName) -> $displayName" -ForegroundColor Green
        }
    }
}

Write-Host "Done." -ForegroundColor Cyan
