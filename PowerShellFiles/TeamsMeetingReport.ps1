
#requires -Version 5.1

# Microsoft Teams Meeting Usage Report
# Scope: Entire Microsoft 365 Tenant
# Reports: D7, D30, D90, D180

$ErrorActionPreference = "Stop"

# ==============================
# Configuration
# ==============================

$Period = "D30"   # D7, D30, D90, D180
$OutputFolder = Join-Path $env:USERPROFILE "Desktop\Teams-Reports"

# ==============================
# Prepare output folder
# ==============================

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force |
        Out-Null
}

$LogFile = Join-Path $OutputFolder "Teams_Report_Errors.log"

# ==============================
# Install and import module
# ==============================

if (-not (Get-Module -ListAvailable Microsoft.Graph.Reports)) {
    Install-Module Microsoft.Graph.Reports `
        -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Reports

# ==============================
# Connect to Microsoft Graph
# ==============================

try {
    Connect-MgGraph -Scopes "Reports.Read.All" `
        -UseDeviceCode -NoWelcome

    $Context = Get-MgContext

    if (-not $Context) {
        throw "Microsoft Graph connection was not established."
    }

    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green
    Write-Host "Tenant: $($Context.TenantId)"
    Write-Host "Period: $Period"
}
catch {
    Add-Content -Path $LogFile -Value (
        "$(Get-Date -Format s) CONNECTION ERROR: $($_.Exception.Message)"
    )

    throw
}

# ==============================
# Download detailed user report
# ==============================

$DetailsFile = Join-Path $OutputFolder "Teams_User_Details.csv"

try {
    Get-MgReportTeamUserActivityUserDetail `
        -Period $Period `
        -OutFile $DetailsFile

    if (-not (Test-Path $DetailsFile)) {
        throw "The report CSV file was not created."
    }

    $Users = @(Import-Csv -Path $DetailsFile)

    Write-Host "Downloaded $($Users.Count) user records." `
        -ForegroundColor Green
}
catch {
    Add-Content -Path $LogFile -Value (
        "$(Get-Date -Format s) REPORT ERROR: $($_.Exception.Message)"
    )

    throw
}

# ==============================
# User ranking
# ==============================

$RankingFile = Join-Path $OutputFolder "Teams_User_Ranking.csv"

$Ranking = $Users |
    Sort-Object {
        [long](
            $_.'Meetings Organized Count'
        )
    } -Descending |
    Select-Object `
        'User Principal Name',
        'Meetings Organized Count',
        'Meetings Attended Count',
        'Meeting Count',
        'Call Count',
        'Audio Duration',
        'Video Duration',
        'Screen Share Duration',
        'Report Period'

$Ranking |
    Export-Csv -Path $RankingFile `
        -NoTypeInformation -Encoding UTF8

# ==============================
# Organization summary
# ==============================

$SummaryFile = Join-Path $OutputFolder "Teams_Report_Summary.csv"

function Convert-ToLong {
    param($Value)

    $Number = 0L

    if ([long]::TryParse(
        [string]$Value,
        [ref]$Number
    )) {
        return $Number
    }

    return 0L
}

$TotalOrganized = 0L
$TotalAttended = 0L
$TotalMeetings = 0L
$TotalCalls = 0L

foreach ($User in $Users) {
    $TotalOrganized += Convert-ToLong `
        $User.'Meetings Organized Count'

    $TotalAttended += Convert-ToLong `
        $User.'Meetings Attended Count'

    $TotalMeetings += Convert-ToLong `
        $User.'Meeting Count'

    $TotalCalls += Convert-ToLong `
        $User.'Call Count'
}

$Summary = [PSCustomObject]@{
    ReportPeriod              = $Period
    ReportRefreshDate         = $Users[0].'Report Refresh Date'
    UsersInReport             = $Users.Count
    TotalMeetingsOrganized    = $TotalOrganized
    TotalMeetingParticipations = $TotalAttended
    TotalMeetingCount         = $TotalMeetings
    TotalCalls                = $TotalCalls
    GeneratedAt               = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
}

$Summary |
    Export-Csv -Path $SummaryFile `
        -NoTypeInformation -Encoding UTF8

# ==============================
# Display results
# ==============================

Write-Host "`n========== TEAMS REPORT ==========" `
    -ForegroundColor Cyan

$Summary | Format-List

Write-Host "Files saved to: $OutputFolder" `
    -ForegroundColor Green

Write-Host "`nTop 10 users by organized meetings:" `
    -ForegroundColor Cyan

$Ranking |
    Select-Object -First 10 |
    Format-Table -AutoSize

Write-Host "`nReport completed successfully." `
    -ForegroundColor Green

#Disconnect-MgGraph
