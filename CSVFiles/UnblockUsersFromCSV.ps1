
# ============================================================
# Unblock Microsoft 365 users from CSV
# Input : C:\SchoolUsers\UsersToUnblock.csv
# Output: C:\SchoolUsers\UnblockUsersReport.csv
# ============================================================

$ErrorActionPreference = "Stop"

$CsvPath    = "C:\SchoolUsers\UsersToUnblock.csv"
$ReportPath = "C:\SchoolUsers\UnblockUsersReport.csv"

# Verify input file
if (-not (Test-Path $CsvPath)) {
    Write-Host "ERROR: CSV file not found: $CsvPath" -ForegroundColor Red
    return
}

# Install required Microsoft Graph modules if needed
foreach ($ModuleName in @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users"
)) {
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        try {
            Install-Module $ModuleName -Scope CurrentUser -Force -ErrorAction Stop
        }
        catch {
            Write-Host "Module installation failed: $ModuleName" -ForegroundColor Red
            Write-Host $_.Exception.Message
            return
        }
    }
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users

# Import and validate CSV
try {
    $Rows = @(Import-Csv -Path $CsvPath -ErrorAction Stop)
}
catch {
    Write-Host "Could not read CSV: $($_.Exception.Message)" -ForegroundColor Red
    return
}

if ($Rows.Count -eq 0 -or
    $Rows[0].PSObject.Properties.Name -notcontains "UserPrincipalName") {
    Write-Host "CSV must contain a UserPrincipalName column." -ForegroundColor Red
    return
}

$Users = @(
    $Rows |
    ForEach-Object {
        if (-not [string]::IsNullOrWhiteSpace($_.UserPrincipalName)) {
            $_.UserPrincipalName.Trim()
        }
    } |
    Sort-Object -Unique
)

if ($Users.Count -eq 0) {
    Write-Host "No valid user accounts found." -ForegroundColor Yellow
    return
}

Write-Host "`nUsers to unblock: $($Users.Count)" -ForegroundColor Cyan
$Users | ForEach-Object { Write-Host " - $_" }

# One confirmation before changes
$Confirmation = Read-Host "`nType YES to enable sign-in for all listed accounts"

if ($Confirmation -cne "YES") {
    Write-Host "Operation cancelled. No changes made." -ForegroundColor Yellow
    return
}

# Device-code authentication to avoid WAM window-handle issues
try {
    Connect-MgGraph `
        -Scopes "User.EnableDisableAccount.All", "User.Read.All" `
        -UseDeviceCode `
        -NoWelcome `
        -ErrorAction Stop
}
catch {
    Write-Host "Microsoft Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
    return
}

$Results = [System.Collections.Generic.List[object]]::new()

foreach ($UPN in $Users) {
    $User = $null
    $Status = "Failed"
    $Details = ""

    Write-Host "`nProcessing: $UPN" -ForegroundColor Cyan

    try {
        $User = Get-MgUser `
            -UserId $UPN `
            -Property "id,displayName,userPrincipalName,accountEnabled" `
            -ErrorAction Stop

        if ($User.AccountEnabled -eq $true) {
            $Status = "AlreadyEnabled"
            $Details = "Sign-in was already enabled"
        }
        else {
            Update-MgUser `
                -UserId $User.Id `
                -AccountEnabled:$true `
                -ErrorAction Stop

            $Status = "Unblocked"
            $Details = "Sign-in enabled successfully"

            Write-Host "Unblocked: $UPN" -ForegroundColor Green
        }
    }
    catch {
        $Details = $_.Exception.Message
        Write-Host "Failed: $UPN - $Details" -ForegroundColor Red
    }

    $Results.Add([PSCustomObject]@{
        UserPrincipalName = $UPN
        DisplayName       = if ($User) { $User.DisplayName } else { "" }
        PreviousState     = if ($User) {
            if ($User.AccountEnabled) { "Enabled" } else { "Disabled" }
        } else { "Unknown" }
        Status            = $Status
        Details           = $Details
        ProcessedAt       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    })
}

# Export results
try {
    $Results | Export-Csv `
        -Path $ReportPath `
        -NoTypeInformation `
        -Encoding UTF8 `
        -ErrorAction Stop

    Write-Host "`nReport saved to: $ReportPath" -ForegroundColor Green
}
catch {
    Write-Host "Could not save report: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n========== SUMMARY ==========" -ForegroundColor Cyan
$Results | Group-Object Status | ForEach-Object {
    Write-Host "$($_.Name): $($_.Count)"
}

Disconnect-MgGraph
