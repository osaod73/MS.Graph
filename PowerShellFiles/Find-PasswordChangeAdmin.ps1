
# ============================================================
# Report: Who changed/reset passwords for users in a CSV file
# Microsoft Graph PowerShell
# ============================================================

$InputCsv  = "C:\SchoolUsers\Users.csv"
$OutputCsv = "C:\SchoolUsers\PasswordChangeAuditReport.csv"

# Number of days of audit history to search
$DaysBack = 10

$ErrorActionPreference = "Stop"

try {
    # Import required Microsoft Graph module
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Reports -ErrorAction Stop

    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "AuditLog.Read.All" `
        -UseDeviceCode -NoWelcome

    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green

    # Validate input file
    if (-not (Test-Path $InputCsv)) {
        throw "Input CSV file not found: $InputCsv"
    }

    $Users = Import-Csv -Path $InputCsv

    if (-not ($Users | Get-Member -Name UserPrincipalName -MemberType NoteProperty)) {
        throw "CSV must contain a UserPrincipalName column."
    }

    # Set audit search start date
    $StartDate = (Get-Date).ToUniversalTime().AddDays(-$DaysBack)
    $FilterDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

    Write-Host "Reading audit logs from the last $DaysBack days..." `
        -ForegroundColor Cyan

    # Retrieve audit records in the selected period
    $AuditLogs = Get-MgAuditLogDirectoryAudit `
        -Filter "activityDateTime ge $FilterDate" `
        -All -ErrorAction Stop

    # Select password-related activities
    $PasswordLogs = $AuditLogs | Where-Object {
        $_.ActivityDisplayName -match
        '(?i)password|credential'
    }

    $Results = foreach ($User in $Users) {

        $UPN = ([string]$User.UserPrincipalName).Trim()

        if ([string]::IsNullOrWhiteSpace($UPN)) {
            continue
        }

        # Match the target user in audit log resources
        $Matches = foreach ($Log in $PasswordLogs) {

            $TargetFound = $false

            foreach ($Target in $Log.TargetResources) {

                $TargetUPN = [string]$Target.UserPrincipalName

                # Some audit records do not populate the UPN field
                if ($TargetUPN -and
                    $TargetUPN.Equals(
                        $UPN,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )) {
                    $TargetFound = $true
                }

                # Also check modified properties where available
                if (-not $TargetFound) {
                    foreach ($Property in $Target.ModifiedProperties) {
                        if ($Property.NewValue -and
                            ([string]$Property.NewValue).Trim('"') -ieq $UPN) {
                            $TargetFound = $true
                        }
                    }
                }
            }

            if ($TargetFound) {
                $Log
            }
        }

        if (-not $Matches) {
            [PSCustomObject]@{
                UserPrincipalName = $UPN
                Activity          = "No matching password event found"
                AdminName         = "Not found"
                AdminUPN          = ""
                DateTime          = ""
                Status            = ""
                InitiatedByType   = ""
                TargetMatched     = ""
            }
            continue
        }

        foreach ($Log in ($Matches | Sort-Object ActivityDateTime -Descending)) {

            $ActorName = ""
            $ActorUPN  = ""
            $ActorType = ""

            if ($Log.InitiatedBy.User) {
                $ActorName = $Log.InitiatedBy.User.DisplayName
                $ActorUPN  = $Log.InitiatedBy.User.UserPrincipalName
                $ActorType = "User"
            }
            elseif ($Log.InitiatedBy.App) {
                $ActorName = $Log.InitiatedBy.App.DisplayName
                $ActorUPN  = $Log.InitiatedBy.App.ServicePrincipalName
                $ActorType = "Application"
            }

            if ([string]::IsNullOrWhiteSpace($ActorName)) {
                $ActorName = "Unknown / System"
            }

            [PSCustomObject]@{
                UserPrincipalName = $UPN
                Activity          = $Log.ActivityDisplayName
                AdminName         = $ActorName
                AdminUPN          = $ActorUPN
                DateTime          = $Log.ActivityDateTime
                Status            = $Log.Result
                InitiatedByType   = $ActorType
                TargetMatched     = $UPN
            }
        }
    }

    # Export report
    $Results |
        Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Report generated successfully." -ForegroundColor Green
    Write-Host "Output: $OutputCsv" -ForegroundColor Yellow
    Write-Host "Records: $(@($Results).Count)" -ForegroundColor Cyan
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
