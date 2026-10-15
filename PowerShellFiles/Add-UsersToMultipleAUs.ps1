
#Requires -Version 5.1

# ==========================================================
# Add Microsoft 365 users to multiple Administrative Units
# Source: One CSV file
# ==========================================================
# -----------------------------


# https://login.microsoft.com/device


$CsvPath = "C:\SchoolUsers\Users-AU.csv"
$LogPath = "C:\SchoolUsers\AU-Assignment-Results.csv"

$ErrorActionPreference = "Stop"

# 1. Install required Graph modules if missing
$RequiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Install-Module $Module -Scope CurrentUser -Force
    }
    Import-Module $Module
}

# 2. Connect to Microsoft Graph
Disconnect-MgGraph -ErrorAction SilentlyContinue

Connect-MgGraph -Scopes @(
    "User.Read.All",
    "AdministrativeUnit.ReadWrite.All"
) -UseDeviceCode -NoWelcome

# Verify connection
Get-MgContext |
    Select-Object Account, TenantId, AuthType, Scopes

# 3. Validate CSV file
if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "CSV file not found: $CsvPath"
}

$Rows = @(Import-Csv -LiteralPath $CsvPath)

if ($Rows.Count -eq 0) {
    throw "CSV file is empty."
}

$RequiredColumns = @(
    "UserPrincipalName",
    "AdministrativeUnits"
)

foreach ($Column in $RequiredColumns) {
    if ($Rows[0].PSObject.Properties.Name -notcontains $Column) {
        throw "Missing CSV column: $Column"
    }
}

# 4. Cache users and Administrative Units
$UserCache = @{}
$AUCache = @{}

# Get all Administrative Units once
$AllAUs = @(Get-MgDirectoryAdministrativeUnit -All)

foreach ($AU in $AllAUs) {
    $AUCache[$AU.Id.ToLowerInvariant()] = $AU

    $NameKey = $AU.DisplayName.ToLowerInvariant()

    if (-not $AUCache.ContainsKey("NAME:$NameKey")) {
        $AUCache["NAME:$NameKey"] = $AU
    }
}

# 5. Process each CSV row
$Results = [System.Collections.Generic.List[object]]::new()

foreach ($Row in $Rows) {

    $UPN = ([string]$Row.UserPrincipalName).Trim()
    $AUValues = ([string]$Row.AdministrativeUnits).Trim()

    if ([string]::IsNullOrWhiteSpace($UPN) -or
        [string]::IsNullOrWhiteSpace($AUValues)) {

        $Results.Add([PSCustomObject]@{
            UserPrincipalName = $UPN
            AdministrativeUnit = $AUValues
            Status = "Skipped"
            Details = "User or AdministrativeUnits is empty"
        })
        continue
    }

    # Resolve user
    try {
        $UserKey = $UPN.ToLowerInvariant()

        if (-not $UserCache.ContainsKey($UserKey)) {
            $User = Get-MgUser `
                -UserId $UPN `
                -Property "id,userPrincipalName"
            $UserCache[$UserKey] = $User
        }

        $User = $UserCache[$UserKey]
    }
    catch {
        $Results.Add([PSCustomObject]@{
            UserPrincipalName = $UPN
            AdministrativeUnit = ""
            Status = "Failed"
            Details = "User lookup failed: $($_.Exception.Message)"
        })
        continue
    }

    # Split AU names or Object IDs using semicolon
    $AUItems = $AUValues -split ";"

    foreach ($AUItem in $AUItems) {

        $AUIdentifier = $AUItem.Trim()

        if ([string]::IsNullOrWhiteSpace($AUIdentifier)) {
            continue
        }

        try {
            # Resolve AU by Object ID first, then exact display name
            $AUKey = $AUIdentifier.ToLowerInvariant()

            if ($AUCache.ContainsKey($AUKey)) {
                $AU = $AUCache[$AUKey]
            }
            elseif ($AUCache.ContainsKey("NAME:$AUKey")) {
                $AU = $AUCache["NAME:$AUKey"]
            }
            else {
                throw "Administrative Unit not found: $AUIdentifier"
            }

            # Check existing membership
            $IsMember = $false

            $Members = Get-MgDirectoryAdministrativeUnitMember `
                -AdministrativeUnitId $AU.Id `
                -All `
                -Property "id"

            foreach ($Member in $Members) {
                if ($Member.Id -eq $User.Id) {
                    $IsMember = $true
                    break
                }
            }

            if ($IsMember) {
                $Status = "AlreadyMember"
                $Details = "User already belongs to this AU"
            }
            else {
                # Add user to AU
                $Body = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($User.Id)"
                }

                New-MgDirectoryAdministrativeUnitMemberByRef `
                    -AdministrativeUnitId $AU.Id `
                    -BodyParameter $Body `
                    -ErrorAction Stop

                $Status = "Success"
                $Details = "User added successfully"
            }

            Write-Host "$Status : $UPN -> $($AU.DisplayName)" `
                -ForegroundColor $(if ($Status -eq "Success") {
                    "Green"
                } else {
                    "Yellow"
                })

            $Results.Add([PSCustomObject]@{
                UserPrincipalName = $UPN
                AdministrativeUnit = $AU.DisplayName
                Status = $Status
                Details = $Details
            })
        }
        catch {
            $Message = $_.Exception.Message

            Write-Warning "$UPN -> $AUIdentifier : $Message"

            $Results.Add([PSCustomObject]@{
                UserPrincipalName = $UPN
                AdministrativeUnit = $AUIdentifier
                Status = "Failed"
                Details = $Message
            })
        }
    }
}

# 6. Export results
$LogDirectory = Split-Path -Parent $LogPath

if ($LogDirectory -and -not (Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force |
        Out-Null
}

$Results | Export-Csv `
    -LiteralPath $LogPath `
    -NoTypeInformation `
    -Encoding UTF8

Write-Host "`nFinished processing." -ForegroundColor Cyan
Write-Host "Results saved to: $LogPath" -ForegroundColor Cyan

#Disconnect-MgGraph
