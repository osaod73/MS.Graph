
#requires -Version 5.1
# https://login.microsoft.com/device

$ErrorActionPreference = "Stop"

# Output folder
$OutputFolder = "C:\SchoolUsers\s\TeamsAUReport"
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

$DateTag = Get-Date -Format "yyyyMMdd_HHmmss"
$FullReport = Join-Path $OutputFolder "AllTeams_$DateTag.csv"
$ReviewReport = Join-Path $OutputFolder "Teams_NeedingReview_$DateTag.csv"
$ErrorReport = Join-Path $OutputFolder "Errors_$DateTag.csv"

$Results = [System.Collections.Generic.List[object]]::new()
$Errors = [System.Collections.Generic.List[object]]::new()

# Install modules if missing
$Modules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Teams",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "Microsoft.Graph.Reports"
)

foreach ($Module in $Modules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Install-Module $Module -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Module -ErrorAction Stop
}

# Connect to Microsoft Graph
$Scopes = @(
    "AdministrativeUnit.Read.All",
    "Group.Read.All",
    "Team.ReadBasic.All",
    "TeamMember.Read.All",
    "Reports.Read.All"
)

Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome


Write-Host "Loading Teams usage report..." -ForegroundColor Cyan

# Usage report for the last 180 days

$ActivityByTeam = @{}

$UsageFile = Join-Path $OutputFolder "UsageTemp.csv"

try {
    Import-Module Microsoft.Graph.Reports -ErrorAction Stop

    Get-MgReportTeamActivityDetail `
        -Period D180 `
        -OutFile $UsageFile `
        -ErrorAction Stop

    if (Test-Path $UsageFile) {
        foreach ($Row in (Import-Csv $UsageFile)) {
            if ($Row.'Team Id') {
                $ActivityByTeam[$Row.'Team Id'] = $Row
            }
        }

        Write-Host "Teams activity report loaded successfully." `
            -ForegroundColor Green
    }
}
catch {
    Write-Warning "Could not retrieve usage report: $($_.Exception.Message)"
}


Write-Host "Loading Administrative Units..." -ForegroundColor Cyan

try {
    $AUs = Get-MgDirectoryAdministrativeUnit -All `
        -Property Id,DisplayName,Description
}
catch {
    throw "Could not read Administrative Units. Check permissions: $($_.Exception.Message)"
}

foreach ($AU in $AUs) {
    Write-Host "Processing AU: $($AU.DisplayName)" -ForegroundColor Yellow

    try {
        $Members = Get-MgDirectoryAdministrativeUnitMember `
            -AdministrativeUnitId $AU.Id -All
    }
    catch {
        $Errors.Add([pscustomobject]@{
            AUName = $AU.DisplayName
            AUId = $AU.Id
            Error = $_.Exception.Message
        })
        continue
    }

    # Get group IDs that are direct members of this AU
    $GroupIds = @(
        $Members |
        Where-Object { $_.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group' } |
        ForEach-Object { $_.Id }
    )

    foreach ($GroupId in $GroupIds) {
        try {
            $Group = Get-MgGroup -GroupId $GroupId `
                -Property Id,DisplayName,CreatedDateTime,Mail,GroupTypes,Visibility

            # A Microsoft 365 group can exist without a Teams team
            $Team = $null
            try {
                $Team = Get-MgTeam -TeamId $GroupId -ErrorAction Stop
            }
            catch {
                continue
            }

            $Owners = @()
            try {
                $Owners = Get-MgGroupOwner -GroupId $GroupId -All
            }
            catch {}

            $OwnerNames = @(
                $Owners | ForEach-Object {
                    if ($_.AdditionalProperties.displayName) {
                        $_.AdditionalProperties.displayName
                    }
                    elseif ($_.AdditionalProperties.userPrincipalName) {
                        $_.AdditionalProperties.userPrincipalName
                    }
                }
            ) -join "; "

            $OwnerUPNs = @(
                $Owners | ForEach-Object {
                    $_.AdditionalProperties.userPrincipalName
                } | Where-Object { $_ }
            ) -join "; "

            $MembersCount = 0
            try {
                $MembersCount = @(
                    Get-MgGroupMember -GroupId $GroupId -All
                ).Count
            }
            catch {}

            $Activity = $ActivityByTeam[$GroupId]

            $LastActivity = $null
            $ActiveUsers = $null
            $ActiveChannels = $null
            $PostMessages = $null
            $ReplyMessages = $null
            $ChannelMessages = $null
            $ActivityStatus = "غير متوفر في تقرير الاستخدام"

            if ($Activity) {
                $LastActivity = $Activity.'Last Activity Date'
                $ActiveUsers = $Activity.'Active users'
                $ActiveChannels = $Activity.'Active Channels'
                $PostMessages = $Activity.'Post Messages'
                $ReplyMessages = $Activity.'Reply Messages'
                $ChannelMessages = $Activity.'Channel Messages'

                $MessageCount = 0
                foreach ($Value in @($PostMessages, $ReplyMessages, $ChannelMessages)) {
                    $ParsedValue = 0
                    if ([int]::TryParse([string]$Value, [ref]$ParsedValue)) {
                        $MessageCount += $ParsedValue
                    }
                }

                if ($LastActivity) {
                    $ActivityStatus = "يوجد نشاط مسجل"
                }
                elseif ($MessageCount -gt 0) {
                    $ActivityStatus = "نشاط رسائل مسجل"
                }
                else {
                    $ActivityStatus = "لا يوجد نشاط ظاهر في الفترة"
                }
            }

            $CreatedDate = $Group.CreatedDateTime
            $AgeDays = $null
            if ($CreatedDate) {
                $AgeDays = [math]::Floor(
                    ((Get-Date).ToUniversalTime() -
                    ([datetime]$CreatedDate).ToUniversalTime()).TotalDays
                )
            }

            $NeedsReview = (
                -not $Activity -or
                -not $LastActivity -or
                ([string]$ActiveUsers -eq "0" -and
                 [string]$ActiveChannels -eq "0")
            )

            $Results.Add([pscustomobject]@{
                AUName = $AU.DisplayName
                AUId = $AU.Id
                TeamName = $Group.DisplayName
                TeamId = $GroupId
                TeamUrl = $Team.WebUrl
                CreatedDateTime = $CreatedDate
                AgeDays = $AgeDays
                Visibility = $Group.Visibility
                IsArchived = $Team.IsArchived
                Owners = $OwnerNames
                OwnerUPNs = $OwnerUPNs
                MemberCount = $MembersCount
                LastActivityDate = $LastActivity
                ActiveUsers = $ActiveUsers
                ActiveChannels = $ActiveChannels
                PostMessages = $PostMessages
                ReplyMessages = $ReplyMessages
                ChannelMessages = $ChannelMessages
                ActivityStatus = $ActivityStatus
                NeedsReview = $NeedsReview
                Assessment = if ($NeedsReview) {
                    "يحتاج إلى مراجعة التنشيط"
                } else {
                    "ظهر نشاط أو بيانات استخدام"
                }
            })
        }
        catch {
            $Errors.Add([pscustomobject]@{
                AUName = $AU.DisplayName
                AUId = $AU.Id
                GroupId = $GroupId
                Error = $_.Exception.Message
            })
        }
    }
}

# Export reports
$Results |
    Sort-Object AUName,TeamName |
    Export-Csv -Path $FullReport -NoTypeInformation -Encoding UTF8

$Results |
    Where-Object { $_.NeedsReview -eq $true } |
    Sort-Object AUName,TeamName |
    Export-Csv -Path $ReviewReport -NoTypeInformation -Encoding UTF8

if ($Errors.Count -gt 0) {
    $Errors | Export-Csv -Path $ErrorReport -NoTypeInformation -Encoding UTF8
}

Write-Host "`nReport completed." -ForegroundColor Green
Write-Host "Total Teams: $($Results.Count)"
Write-Host "Teams needing review: $( @($Results | Where-Object NeedsReview).Count )"
Write-Host "Full report: $FullReport"
Write-Host "Review report: $ReviewReport"

if ($Errors.Count -gt 0) {
    Write-Host "Error report: $ErrorReport" -ForegroundColor Red
}

Disconnect-MgGraph
