
# ==========================================================
# Bulk Password Reset - Microsoft 365 / Microsoft Graph
# Force users to change password at first sign-in
# ==========================================================

$ErrorActionPreference = "Stop"

# CSV file path
$CsvPath = "C:\SchoolUsers\UsersPasswords.csv"

# Report file
$ReportPath = "C:\SchoolUsers\PasswordChangeReport.csv"

# ----------------------------------------------------------
# 1. Install and import Microsoft Graph module
# ----------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Users

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Authentication

# ----------------------------------------------------------
# 2. Connect to Microsoft Graph
# ----------------------------------------------------------

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

try {
    Connect-MgGraph -Scopes "User.ReadWrite.All" -ContextScope Process -NoWelcome
}
catch {
    Write-Host "Connection failed. Trying device code authentication..." -ForegroundColor Yellow
    $env:MSAL_USE_WAM = "false"
    Connect-MgGraph -Scopes "User.ReadWrite.All" -UseDeviceCode -ContextScope Process -NoWelcome
}

$Context = Get-MgContext

if (-not $Context) {
    Write-Host "Microsoft Graph connection failed!" -ForegroundColor Red
    exit
}

Write-Host "Connected successfully!" -ForegroundColor Green
Write-Host "Tenant ID: $($Context.TenantId)" -ForegroundColor Cyan

# ----------------------------------------------------------
# 3. Read CSV
# ----------------------------------------------------------

if (-not (Test-Path $CsvPath)) {
    Write-Host "CSV file not found: $CsvPath" -ForegroundColor Red
    exit
}

$Users = Import-Csv -Path $CsvPath

if (-not $Users -or
    -not ($Users[0].PSObject.Properties.Name -contains "UserPrincipalName") -or
    -not ($Users[0].PSObject.Properties.Name -contains "NewPassword")) {

    Write-Host "Invalid CSV format!" -ForegroundColor Red
    exit
}

# ----------------------------------------------------------
# 4. Confirmation
# ----------------------------------------------------------

Write-Host ""
Write-Host "Users to process: $($Users.Count)" -ForegroundColor Yellow
Write-Host "All users will be required to change their password at first sign-in." -ForegroundColor Yellow

$Confirm = Read-Host "Type YES to continue"

if ($Confirm -cne "YES") {
    Write-Host "Operation cancelled." -ForegroundColor Red
    exit
}

# ----------------------------------------------------------
# 5. Change passwords
# ----------------------------------------------------------

$Results = @()
$Counter = 0

foreach ($User in $Users) {

    $Counter++

    $UPN = ([string]$User.UserPrincipalName).Trim()
    $Password = [string]$User.NewPassword

    Write-Host ""
    Write-Host "[$Counter/$($Users.Count)] Processing: $UPN" -ForegroundColor Cyan

    $Status = "Failed"
    $Message = ""

    try {

        if ([string]::IsNullOrWhiteSpace($UPN) -or
            [string]::IsNullOrEmpty($Password)) {
            throw "Missing username or password."
        }

        # Password profile
        $PasswordProfile = @{
            Password = $Password
            ForceChangePasswordNextSignIn = $true
        }

        # Update user password
        Update-MgUser `
            -UserId $UPN `
            -PasswordProfile $PasswordProfile `
            -ErrorAction Stop

        $Status = "Success"
        $Message = "Password changed. User must change it at first sign-in."

        Write-Host "SUCCESS: $UPN" -ForegroundColor Green
    }
    catch {

        $Message = $_.Exception.Message

        Write-Host "FAILED: $UPN" -ForegroundColor Red
        Write-Host $Message -ForegroundColor Yellow
    }

    # Record result without storing passwords
    $Results += [PSCustomObject]@{
        UserPrincipalName = $UPN
        Status            = $Status
        ForceChange       = $true
        Date              = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Message           = $Message
    }
}

# ----------------------------------------------------------
# 6. Export report
# ----------------------------------------------------------

$Results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

# ----------------------------------------------------------
# 7. Summary
# ----------------------------------------------------------

$SuccessCount = @($Results | Where-Object Status -eq "Success").Count
$FailedCount = @($Results | Where-Object Status -eq "Failed").Count

Write-Host ""
Write-Host "====================================" -ForegroundColor Cyan
Write-Host " PASSWORD RESET SUMMARY" -ForegroundColor Cyan
Write-Host "====================================" -ForegroundColor Cyan

Write-Host "Total Users  : $($Users.Count)"
Write-Host "Successful   : $SuccessCount" -ForegroundColor Green
Write-Host "Failed       : $FailedCount" -ForegroundColor Red

Write-Host ""
Write-Host "Report saved to:" -ForegroundColor Yellow
Write-Host $ReportPath -ForegroundColor Green

Write-Host ""
Write-Host "Operation completed." -ForegroundColor Cyan
