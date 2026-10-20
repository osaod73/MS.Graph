
# ============================================================
# Report users with no recorded sign-in attempts in last 30 days
# Includes Department and JobTitle
# Microsoft 365 / Microsoft Entra ID
# ============================================================

$ErrorActionPreference = "Stop"

# Install Microsoft Graph module if not installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Users

# Connect to Microsoft Graph
try {
    Connect-MgGraph -Scopes @(
        "User.Read.All",
        "AuditLog.Read.All"
    ) -NoWelcome -ContextScope Process -UseDeviceAuthentication
}
catch {
    Write-Host "Authentication failed: $($_.Exception.Message)" -ForegroundColor Red
    exit
}

# Define the last 30 days cutoff in UTC
$CutoffDate = (Get-Date).ToUniversalTime().AddDays(-30)

# Output CSV file on Desktop
$OutputPath = Join-Path $env:USERPROFILE "Desktop\InactiveUsersLast30Days.csv"

Write-Host "Retrieving Microsoft 365 users..." -ForegroundColor Cyan

try {
    # Retrieve all users with required properties
    $Users = Get-MgUser -All -Property @(
        "DisplayName",
        "UserPrincipalName",
        "Mail",
        "Department",
        "JobTitle",
        "AccountEnabled",
        "UserType",
        "SignInActivity"
    )

    $Report = foreach ($User in $Users) {

        $Activity = $User.SignInActivity

        $LastInteractive = $null
        $LastNonInteractive = $null

        if ($null -ne $Activity) {
            $LastInteractive = $Activity.LastSignInDateTime
            $LastNonInteractive = $Activity.LastNonInteractiveSignInDateTime
        }

        # Convert timestamps to UTC
        $InteractiveDate = $null
        $NonInteractiveDate = $null

        if (-not [string]::IsNullOrWhiteSpace("$LastInteractive")) {
            $InteractiveDate = ([DateTimeOffset]$LastInteractive).UtcDateTime
        }

        if (-not [string]::IsNullOrWhiteSpace("$LastNonInteractive")) {
            $NonInteractiveDate = ([DateTimeOffset]$LastNonInteractive).UtcDateTime
        }

        # Determine the latest recorded sign-in attempt
        $Dates = @(
            $InteractiveDate
            $NonInteractiveDate
        ) | Where-Object { $null -ne $_ }

        $LastAttempt = $null

        if (@($Dates).Count -gt 0) {
            $LastAttempt = ($Dates | Measure-Object -Maximum).Maximum
        }

        # Include users with no recorded attempt or whose
        # latest recorded attempt was more than 30 days ago
        if (($null -eq $LastAttempt) -or ($LastAttempt -lt $CutoffDate)) {

            if ($null -eq $LastAttempt) {
                $Status = "NoRecordedSignIn"
            }
            else {
                $Status = "NoSignInAttemptsInLast30Days"
            }

            [PSCustomObject]@{
                DisplayName              = $User.DisplayName
                UserPrincipalName        = $User.UserPrincipalName
                Mail                     = $User.Mail
                Department               = $User.Department
                JobTitle                 = $User.JobTitle
                AccountEnabled           = $User.AccountEnabled
                UserType                 = $User.UserType
                LastInteractiveSignIn    = $LastInteractive
                LastNonInteractiveSignIn = $LastNonInteractive
                LastRecordedAttempt      = $LastAttempt
                Status                   = $Status
            }
        }
    }

    # Export report to CSV
    $Report |
        Sort-Object Department, DisplayName |
        Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    # Display summary
    Write-Host ""
    Write-Host "Report generated successfully." -ForegroundColor Green
    Write-Host "Total users in report: $(@($Report).Count)" -ForegroundColor Yellow
    Write-Host "Report location: $OutputPath" -ForegroundColor Cyan

    # Display selected columns in PowerShell
    $Report |
        Select-Object DisplayName,
                      UserPrincipalName,
                      Department,
                      JobTitle,
                      LastRecordedAttempt,
                      Status |
        Format-Table -AutoSize
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
