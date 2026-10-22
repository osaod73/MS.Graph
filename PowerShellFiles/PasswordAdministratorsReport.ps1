
# ============================================================
# Report: Password Administrators in All Administrative Units
# Microsoft Entra ID / Microsoft Graph PowerShell
# ============================================================

$ErrorActionPreference = "Stop"

$ReportPath = Join-Path $env:USERPROFILE "Desktop\PasswordAdministratorsReport.csv"

# Install required modules if missing
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}

if (-not (Get-Module -ListAvailable Microsoft.Graph.Identity.DirectoryManagement)) {
    Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force
}

if (-not (Get-Module -ListAvailable Microsoft.Graph.Identity.Governance)) {
    Install-Module Microsoft.Graph.Identity.Governance -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Identity.DirectoryManagement
Import-Module Microsoft.Graph.Identity.Governance

# Connect to Microsoft Graph
# Use device-code authentication to avoid WAM window-handle errors.
Connect-MgGraph -Scopes @(
    "Directory.Read.All",
    "RoleManagement.Read.Directory",
    "User.Read.All",
    "Group.Read.All"
) -UseDeviceCode -NoWelcome

try {
    Write-Host "Reading administrative units..." -ForegroundColor Cyan

    $units = Get-MgDirectoryAdministrativeUnit -All

    Write-Host "Reading role definitions..." -ForegroundColor Cyan

    $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All
    $roleMap = @{}

    foreach ($role in $roleDefinitions) {
        $roleMap[$role.Id] = $role.DisplayName
    }

    # Roles that can provide password reset or user password management.
    $passwordRoles = @(
        "Password Administrator",
        "Authentication Administrator",
        "Privileged Authentication Administrator",
        "User Administrator",
        "Global Administrator",
        "Helpdesk Administrator"
    )

    Write-Host "Reading role assignments..." -ForegroundColor Cyan

    $assignments = Get-MgRoleManagementDirectoryRoleAssignment -All

    # Cache users and groups to avoid repeated lookups.
    $principalCache = @{}
    $report = [System.Collections.Generic.List[object]]::new()

    # Map AU IDs to names
    $unitMap = @{}
    foreach ($unit in $units) {
        $unitMap[$unit.Id] = $unit.DisplayName
    }

    foreach ($assignment in $assignments) {

        $roleName = $roleMap[$assignment.RoleDefinitionId]

        if ($roleName -notin $passwordRoles) {
            continue
        }

        $scopeId = $assignment.DirectoryScopeId

        # Include tenant-wide roles and AU-scoped roles only.
        if ($scopeId -eq "/") {
            $scopeType = "Tenant-wide"
            $unitName = "All Administrative Units"
        }
        elseif ($scopeId -match "^/administrativeUnits/([^/]+)$") {
            $auId = $Matches[1]

            if (-not $unitMap.ContainsKey($auId)) {
                continue
            }

            $scopeType = "Administrative Unit"
            $unitName = $unitMap[$auId]
        }
        else {
            continue
        }

        $principalId = $assignment.PrincipalId

        if (-not $principalCache.ContainsKey($principalId)) {
            try {
                $uri = "https://graph.microsoft.com/v1.0/directoryObjects/$principalId"
                $obj = Invoke-MgGraphRequest -Method GET -Uri $uri

                $principalCache[$principalId] = $obj
            }
            catch {
                $principalCache[$principalId] = $null
            }
        }

        $principal = $principalCache[$principalId]

        if ($null -eq $principal) {
            $principalType = "Unknown"
            $displayName = "Unresolved principal"
            $upn = ""
            $email = ""
        }
        else {
            $odataType = $principal.'@odata.type'

            if ($odataType -eq "#microsoft.graph.user") {
                $principalType = "User"
                $displayName = $principal.displayName
                $upn = $principal.userPrincipalName
                $email = $principal.mail
            }
            elseif ($odataType -eq "#microsoft.graph.group") {
                $principalType = "Group"
                $displayName = $principal.displayName
                $upn = ""
                $email = ""
            }
            else {
                $principalType = $odataType
                $displayName = $principal.displayName
                $upn = ""
                $email = ""
            }
        }

        $report.Add([PSCustomObject]@{
            AdministrativeUnit = $unitName
            ScopeType          = $scopeType
            PrincipalName      = $displayName
            PrincipalType      = $principalType
            UserPrincipalName  = $upn
            Email              = $email
            RoleName           = $roleName
            PrincipalId        = $principalId
            DirectoryScopeId   = $scopeId
        })
    }

    $report |
        Sort-Object AdministrativeUnit, RoleName, PrincipalName |
        Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Report completed successfully." -ForegroundColor Green
    Write-Host "Total records: $($report.Count)"
    Write-Host "Report saved to: $ReportPath" -ForegroundColor Yellow

    # Display report
    $report |
        Format-Table AdministrativeUnit, PrincipalName,
                     PrincipalType, RoleName, ScopeType -AutoSize
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
