
# ============================================================
# Block Microsoft 365 users from CSV
# Input : C:\SchoolUsers\UsersToBlock.csv
# Output: C:\SchoolUsers\BlockUsersReport.csv
# ============================================================

$ErrorActionPreference = "Stop"

$CsvPath    = "C:\SchoolUsers\UsersToBlock.csv"
$ReportPath = "C:\SchoolUsers\BlockUsersReport.csv"

# Check CSV file
if (-not (Test-Path $CsvPath)) {
    Write-Host "ERROR: CSV file not found: $CsvPath" -ForegroundColor Red
    return
}

# Install required module if missing
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users

# Read and validate CSV
$Users = @(Import-Csv -Path $CsvPath)

if ($Users.Count -eq 0 -or
    $Users[0].PSObject.Properties.Name -notcontains "UserPrincipalName") {
    Write-Host "ERROR: CSV must contain UserPrincipalName column." -ForegroundColor Red
    return
}

$Users = @(
    $Users |
    Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName)
    } |
    ForEach-Object {
        $_.UserPrincipalName.Trim()
    } |
    Sort-Object -Unique
)

if ($Users.Count -eq 0) {
    Write-Host "No valid user accounts found." -ForegroundColor Yellow
    return
}

Write-Host "`nAccounts to block: $($Users.Count)" -ForegroundColor Cyan
$Users | ForEach-Object { Write-Host " - $_" }

# One confirmation before any changes
$Confirmation = Read-Host "`nType YES to block all listed accounts"

if ($Confirmation -cne "YES") {
    Write-Host "Operation cancelled. No accounts were changed." -ForegroundColor Yellow
    return
}

# Connect using device code to avoid the WAM window-handle issue
try {
    Connect-MgGraph `
        -Scopes "User.EnableDisableAccount.All", "User.Read.All" `
        -UseDeviceCode `
        -NoWelcome `
        -ErrorAction Stop
}
catch {
    Write-Host "Graph connection failed: $($_.Exception.Message)" -ForegroundColor Red
    return
}

$Results = [System.Collections.Generic.List[object]]::new()

foreach ($UPN in $Users) {
    Write-Host "`nProcessing: $UPN" -ForegroundColor Cyan

    try {
        $User = Get-MgUser `
            -UserId $UPN `
            -Property "id,userPrincipalName,displayName,accountEnabled" `
            -ErrorAction Stop

        if ($User.AccountEnabled -eq $false) {
            $Status = "AlreadyBlocked"
            $Details = "Account was already disabled"
        }
        else {
            Update-MgUser `
                -UserId $User.Id `
                -AccountEnabled:$false `
                -ErrorAction Stop

            $Status = "Blocked"
            $Details = "Sign-in disabled successfully"

            Write-Host "Blocked: $UPN" -ForegroundColor Green
        }
    }
    catch {
        $Status = "Failed"
        $Details = $_.Exception.Message

        Write-Host "Failed: $UPN - $Details" -ForegroundColor Red
    }

    $Results.Add([PSCustomObject]@{
        UserPrincipalName = $UPN
        DisplayName       = if ($User) { $User.DisplayName } else { "" }
        Status            = $Status
        Details           = $Details
        ProcessedAt       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    })

    $User = $null
}

# Export report
$Results | Export-Csv `
    -Path $ReportPath `
    -NoTypeInformation `
    -Encoding UTF8

Write-Host "`n========== SUMMARY ==========" -ForegroundColor Cyan
$Results | Group-Object Status | ForEach-Object {
    Write-Host "$($_.Name): $($_.Count)"
}

Write-Host "`nReport saved to: $ReportPath" -ForegroundColor Green

Disconnect-MgGraph
