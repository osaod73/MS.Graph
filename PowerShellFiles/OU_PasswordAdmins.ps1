
# ============================================================
# Microsoft Entra ID - Administrative Unit Password Admins
#
# CSV columns:
# ObjectId,userPrincipalName
#
# Assigns the Password Administrator role at AU scope.
# Existing role assignments are preserved.
# ============================================================

$ErrorActionPreference = "Stop"

# File paths
$CsvPath = "C:\SchoolUsers\OU_PasswordAdmins.csv"
$LogPath = "C:\SchoolUsers\OU_PasswordAdmins_Log.csv"

# Required Microsoft Graph modules
$RequiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "Microsoft.Graph.Identity.Governance"
)

foreach ($ModuleName in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        throw "Module '$ModuleName' is missing. Install it using Install-Module '$ModuleName' -Scope CurrentUser"
    }

    Import-Module $ModuleName -ErrorAction Stop
}

# Validate CSV file
if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "CSV file not found: $CsvPath"
}

$Rows = @(Import-Csv -LiteralPath $CsvPath)

if ($Rows.Count -eq 0) {
    throw "The CSV file is empty."
}

$Headers = @($Rows[0].PSObject.Properties.Name)

if (
    $Headers -notcontains "ObjectId" -or
    $Headers -notcontains "userPrincipalName"
) {
    throw "CSV headers must be ObjectId and userPrincipalName."
}

# Create log directory if necessary
$LogDirectory = Split-Path -Parent $LogPath

if (-not (Test-Path -LiteralPath $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force |
        Out-Null
}

# Connect to Microsoft Graph
$Scopes = @(
    "User.Read.All",
    "AdministrativeUnit.Read.All",
    "RoleManagement.Read.Directory",
    "RoleManagement.ReadWrite.Directory"
)

Connect-MgGraph -Scopes $Scopes -UseDeviceCode -ErrorAction Stop

# Find the Password Administrator role definition
$Roles = @(
    Get-MgRoleManagementDirectoryRoleDefinition `
        -Filter "displayName eq 'Password Administrator'" `
        -All
)

if ($Roles.Count -ne 1) {
    throw "Could not uniquely identify the Password Administrator role."
}

$Role = $Roles[0]

Write-Host "Role: $($Role.DisplayName)" -ForegroundColor Cyan
Write-Host "Processing $($Rows.Count) CSV entries..." -ForegroundColor Cyan

$Results = foreach ($Row in $Rows) {

    $ObjectId = ([string]$Row.ObjectId).Trim()
    $UPN = ([string]$Row.userPrincipalName).Trim()

    $UnitName = ""
    $Status = ""

    try {
        # Validate input
        if (
            [string]::IsNullOrWhiteSpace($ObjectId) -or
            [string]::IsNullOrWhiteSpace($UPN)
        ) {
            throw "ObjectId or userPrincipalName is empty."
        }

        # Retrieve Administrative Unit using Object ID
        $Unit = Get-MgDirectoryAdministrativeUnit `
            -AdministrativeUnitId $ObjectId `
            -ErrorAction Stop

        $UnitName = $Unit.DisplayName

        # Retrieve user using User Principal Name
        $EscapedUPN = $UPN.Replace("'", "''")

        $Users = @(
            Get-MgUser `
                -Filter "userPrincipalName eq '$EscapedUPN'" `
                -Property Id,DisplayName,UserPrincipalName `
                -All
        )

        if ($Users.Count -ne 1) {
            throw "Expected one user for '$UPN'; found $($Users.Count)."
        }

        $User = $Users[0]

        # AU-scoped role assignment
        $Scope = "/administrativeUnits/$($Unit.Id)"

        # Check existing assignments for this user
        $Assignments = @(
            Get-MgRoleManagementDirectoryRoleAssignment `
                -Filter "principalId eq '$($User.Id)'" `
                -All
        )

        $Existing = @(
            $Assignments | Where-Object {
                $_.RoleDefinitionId -eq $Role.Id -and
                $_.DirectoryScopeId -eq $Scope
            }
        )

        if ($Existing.Count -gt 0) {
            $Status = "SKIPPED - assignment already exists"
        }
        else {
            # Create the scoped role assignment
            $Body = @{
                principalId      = $User.Id
                roleDefinitionId = $Role.Id
                directoryScopeId = $Scope
            }

            New-MgRoleManagementDirectoryRoleAssignment `
                -BodyParameter $Body `
                -ErrorAction Stop | Out-Null

            $Status = "SUCCESS - role assigned"
        }

        Write-Host "$Status | $UPN | $UnitName" `
            -ForegroundColor Green
    }
    catch {
        $Status = "ERROR - $($_.Exception.Message)"

        Write-Warning "$UPN | $ObjectId | $Status"
    }

    # Return a log record for every CSV row
    [PSCustomObject]@{
        ObjectId          = $ObjectId
        AdministrativeUnit = $UnitName
        UserPrincipalName = $UPN
        Role              = $Role.DisplayName
        Status            = $Status
        ProcessedAt       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
}

# Export results
$Results | Export-Csv `
    -LiteralPath $LogPath `
    -NoTypeInformation `
    -Encoding UTF8

Write-Host ""
Write-Host "Processing completed." -ForegroundColor Cyan
Write-Host "Log file: $LogPath" -ForegroundColor Cyan

$Results | Format-Table `
    ObjectId, AdministrativeUnit, UserPrincipalName, Status `
    -AutoSize
