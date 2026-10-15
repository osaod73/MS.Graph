
# ============================================================
# Microsoft 365 - Users inactive for 365 days
# Read-only report using Microsoft Graph
# ============================================================

$ErrorActionPreference = "Stop"

# Report settings
$DaysInactive = 365
$CutoffDate = (Get-Date).ToUniversalTime().AddDays(-$DaysInactive)

$OutputFolder = Join-Path $env:USERPROFILE "Desktop"
$CsvPath = Join-Path $OutputFolder "Inactive-M365-Users.csv"

try {
    # Install required module if missing
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
        Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
    }

    Import-Module Microsoft.Graph.Users

    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "User.Read.All", "AuditLog.Read.All" `
        -UseDeviceCode -NoWelcome

    Write-Host "Retrieving Microsoft 365 users..." -ForegroundColor Cyan

    # Retrieve all users and sign-in activity
    $Users = Get-MgUser -All -Property @(
        "Id",
        "DisplayName",
        "UserPrincipalName",
        "Mail",
        "AccountEnabled",
        "UserType",
        "CreatedDateTime",
        "SignInActivity"
    )

    $Results = foreach ($User in $Users) {

        $LastSuccessful = $User.SignInActivity.LastSuccessfulSignInDateTime
        $LastInteractive = $User.SignInActivity.LastSignInDateTime
        $LastNonInteractive = $User.SignInActivity.LastNonInteractiveSignInDateTime

        # Use successful sign-in as the main inactivity criterion.
        # Blank means no successful sign-in date is available.
        $Inactive = $false
        $Status = ""

        if ([string]::IsNullOrWhiteSpace([string]$LastSuccessful)) {
            $Inactive = $true
            $Status = "No successful sign-in date recorded"
        }
        elseif ([datetimeoffset]$LastSuccessful -lt $CutoffDate) {
            $Inactive = $true
            $Status = "Inactive for 365 days or more"
        }

        if ($Inactive) {
            [PSCustomObject]@{
                DisplayName                    = $User.DisplayName
                UserPrincipalName              = $User.UserPrincipalName
                Email                          = $User.Mail
                AccountEnabled                 = $User.AccountEnabled
                UserType                       = $User.UserType
                CreatedDateTime                = $User.CreatedDateTime
                LastSuccessfulSignInDateTime   = $LastSuccessful
                LastInteractiveSignInDateTime  = $LastInteractive
                LastNonInteractiveSignInDateTime = $LastNonInteractive
                Status                         = $Status
                ReportDateUTC                  = (Get-Date).ToUniversalTime().ToString("o")
            }
        }
    }

    # Export report
    $Results |
        Sort-Object LastSuccessfulSignInDateTime |
        Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Report completed successfully." -ForegroundColor Green
    Write-Host "Total users scanned: $($Users.Count)"
    Write-Host "Inactive / no successful sign-in date: $(@($Results).Count)"
    Write-Host "CSV file: $CsvPath" -ForegroundColor Yellow

    # Display report
    $Results |
        Select-Object DisplayName, UserPrincipalName,
            LastSuccessfulSignInDateTime, AccountEnabled, Status |
        Format-Table -AutoSize
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
