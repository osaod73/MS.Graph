
# ============================================================
# Report: Users with the most password changes in last 30 days
# Microsoft 365 / Microsoft Graph
# ============================================================

$ErrorActionPreference = "Stop"

# Configuration
$DaysBack = 30
$TopUsers = 100
$OutputFolder = Join-Path $env:USERPROFILE "Documents"

$StartDate = (Get-Date).ToUniversalTime().AddDays(-$DaysBack)
$EndDate = (Get-Date).ToUniversalTime()

$SummaryFile = Join-Path $OutputFolder "TopPasswordChanges.csv"
$DetailsFile = Join-Path $OutputFolder "PasswordChangeDetails.csv"

try {
    # Install Microsoft Graph module if needed
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
    }

    # Connect to Microsoft Graph
    Import-Module Microsoft.Graph.Authentication

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

   Connect-MgGraph -Scopes "AuditLog.Read.All" -UseDeviceCode -NoWelcome

    # Query audit events in the requested period
    $FilterDate = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    $Filter = "activityDateTime ge $FilterDate"

    Write-Host "Retrieving audit events..." -ForegroundColor Cyan

    $Events = Get-MgAuditLogDirectoryAudit `
        -Filter $Filter `
        -All `
        -Property "id,activityDateTime,activityDisplayName,result,initiatedBy,targetResources,loggedByService,category"

    # Keep successful password-related events only
    $PasswordEvents = $Events | Where-Object {
        $_.Result -eq "success" -and
        (
            $_.ActivityDisplayName -match "(?i)password" -or
            $_.LoggedByService -match "(?i)password management"
        )
    }

    $Details = foreach ($Event in $PasswordEvents) {
        $Targets = @(
            $Event.TargetResources | Where-Object {
                $_.Type -eq "User" -or $_.UserPrincipalName
            }
        )

        foreach ($Target in $Targets) {
            $UserUPN = $Target.UserPrincipalName

            if ([string]::IsNullOrWhiteSpace($UserUPN)) {
                $UserUPN = $Target.DisplayName
            }

            if ([string]::IsNullOrWhiteSpace($UserUPN)) {
                continue
            }

            $ActorName = $null
            $ActorUPN = $null

            if ($Event.InitiatedBy.User) {
                $ActorName = $Event.InitiatedBy.User.DisplayName
                $ActorUPN = $Event.InitiatedBy.User.UserPrincipalName
            }
            elseif ($Event.InitiatedBy.App) {
                $ActorName = $Event.InitiatedBy.App.DisplayName
                $ActorUPN = $Event.InitiatedBy.App.AppId
            }

            [PSCustomObject]@{
                UserPrincipalName = $UserUPN
                UserDisplayName   = $Target.DisplayName
                Activity          = $Event.ActivityDisplayName
                DateTimeUTC       = $Event.ActivityDateTime
                InitiatedBy       = $ActorName
                InitiatorUPN      = $ActorUPN
                Result            = $Event.Result
                Service           = $Event.LoggedByService
                EventId           = $Event.Id
            }
        }
    }

    # Export detailed events
    $Details |
        Sort-Object DateTimeUTC -Descending |
        Export-Csv -Path $DetailsFile -NoTypeInformation -Encoding UTF8

    # Group events by user and rank by count
    $Summary = $Details |
        Group-Object UserPrincipalName |
        ForEach-Object {
            $UserEvents = @($_.Group)

            [PSCustomObject]@{
                UserPrincipalName = $_.Name
                PasswordEventCount = $UserEvents.Count
                LastPasswordEvent = (
                    $UserEvents |
                    Sort-Object DateTimeUTC -Descending |
                    Select-Object -First 1
                ).DateTimeUTC
                Activities = (
                    $UserEvents.Activity |
                    Sort-Object -Unique
                ) -join "; "
            }
        } |
        Sort-Object PasswordEventCount -Descending |
        Select-Object -First $TopUsers

    $Summary |
        Export-Csv -Path $SummaryFile -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Top users by password-related audit events:" -ForegroundColor Green

    $Summary |
        Format-Table UserPrincipalName, PasswordEventCount, LastPasswordEvent -AutoSize

    Write-Host ""
    Write-Host "Summary report: $SummaryFile" -ForegroundColor Yellow
    Write-Host "Detailed report: $DetailsFile" -ForegroundColor Yellow
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
