
# ============================================================
# Restore Deleted Microsoft 365 Groups - Last 72 Hours
# Microsoft Graph | PowerShell
# ============================================================

$ErrorActionPreference = "Stop"

# ---------------- CONFIGURATION ----------------

# Enter your Microsoft 365 Tenant ID or tenant domain
$TenantId = "YOUR-TENANT-ID-OR-DOMAIN"

$HoursBack = 72
$ExecuteRestore = $false  # $false = Preview, $true = Restore

$ReportPath = "C:\SchoolUsers\RestoreGroupsLast72Hours.csv"

# ---------------- PREPARE REPORT ----------------

$ReportFolder = Split-Path -Path $ReportPath -Parent

if (-not (Test-Path -LiteralPath $ReportFolder)) {
    New-Item -Path $ReportFolder -ItemType Directory -Force |
        Out-Null
}

$Results = [System.Collections.Generic.List[object]]::new()

# ---------------- LOAD MODULE ----------------

try {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Install-Module Microsoft.Graph.Authentication `
            -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}
catch {
    Write-Host "MODULE ERROR: $($_.Exception.Message)" -ForegroundColor Red
    return
}

# ---------------- CONNECT TO GRAPH ----------------

try {
    if ([string]::IsNullOrWhiteSpace($TenantId) -or
        $TenantId -eq "YOUR-TENANT-ID-OR-DOMAIN") {

        $TenantId = Read-Host "Enter your Microsoft 365 Tenant ID or tenant domain"
    }

    Connect-MgGraph `
        -TenantId $TenantId `
        -Scopes "Group.ReadWrite.All" `
        -UseDeviceCode `
        -NoWelcome `
        -ErrorAction Stop

    $Context = Get-MgContext

    if (-not $Context -or -not $Context.TenantId) {
        throw "Microsoft Graph connection was not established."
    }

    Write-Host "Connected to Microsoft Graph successfully." `
        -ForegroundColor Green

    Write-Host "Tenant ID: $($Context.TenantId)"
}
catch {
    Write-Host "CONNECTION ERROR: $($_.Exception.Message)" `
        -ForegroundColor Red
    return
}

# Do not call Disconnect-MgGraph here.
# This avoids the MSAL cache warning seen in this environment.

# ---------------- DEFINE TIME WINDOW ----------------

$NowUtc = [datetime]::UtcNow
$CutoffUtc = $NowUtc.AddHours(-$HoursBack)

Write-Host ""
Write-Host "Search window: $CutoffUtc UTC to $NowUtc UTC"
Write-Host "Mode: $(if ($ExecuteRestore) { 'RESTORE' } else { 'PREVIEW ONLY' })"
Write-Host ""

# ---------------- READ DELETED GROUPS ----------------

$AllGroups = [System.Collections.Generic.List[object]]::new()

try {
    $Uri = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group?`$select=id,displayName,description,mail,mailNickname,groupTypes,deletedDateTime&`$top=999"

    while ($Uri) {
        $Page = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Uri `
            -OutputType PSObject `
            -ErrorAction Stop

        foreach ($Group in @($Page.value)) {
            if ($null -ne $Group -and $Group.id) {
                $AllGroups.Add($Group)
            }
        }

        $Uri = $Page.'@odata.nextLink'
    }
}
catch {
    Write-Host "GROUP READ ERROR: $($_.Exception.Message)" `
        -ForegroundColor Red

    Write-Host "Check Graph permissions and tenant access."
    return
}

# ---------------- FILTER LAST 72 HOURS ----------------

$Targets = @(
    foreach ($Group in $AllGroups) {
        if (-not $Group.deletedDateTime) {
            continue
        }

        try {
            $DeletedAt = [datetimeoffset]::Parse(
                [string]$Group.deletedDateTime
            ).UtcDateTime

            if ($DeletedAt -ge $CutoffUtc -and
                $DeletedAt -le $NowUtc) {
                $Group
            }
        }
        catch {
            Write-Host "Invalid deletion date for group ID $($Group.id)" `
                -ForegroundColor Yellow
        }
    }
)

Write-Host "Total deleted groups retrieved: $($AllGroups.Count)"
Write-Host "Groups deleted in the last $HoursBack hours: $($Targets.Count)" `
    -ForegroundColor Cyan

if ($Targets.Count -eq 0) {
    Write-Host "No matching groups found."
}

# ---------------- RESTORE OR PREVIEW ----------------

foreach ($Group in $Targets) {

    $DeletedAt = [datetimeoffset]::Parse(
        [string]$Group.deletedDateTime
    ).UtcDateTime

    $Status = "PreviewOnly"
    $Message = "Preview only. No changes made."

    Write-Host ""
    Write-Host "Group: $($Group.displayName)"
    Write-Host "ID: $($Group.id)"
    Write-Host "Deleted: $($DeletedAt.ToString('yyyy-MM-dd HH:mm:ss')) UTC"

    if ($ExecuteRestore) {
        try {
            # Correct Microsoft Graph restore endpoint
            $RestoreUri = "https://graph.microsoft.com/v1.0/directory/deletedItems/$($Group.id)/restore"

            $RestoreResponse = Invoke-MgGraphRequest `
                -Method POST `
                -Uri $RestoreUri `
                -OutputType PSObject `
                -ErrorAction Stop

            $Status = "Restored"
            $Message = "Restore request succeeded."

            Write-Host "RESTORED SUCCESSFULLY" -ForegroundColor Green
        }
        catch {
            $Status = "Failed"
            $Message = $_.Exception.Message

            Write-Host "RESTORE FAILED: $Message" -ForegroundColor Red
        }
    }
    else {
        Write-Host "PREVIEW ONLY - not restored." -ForegroundColor Yellow
    }

    $Results.Add([PSCustomObject]@{
        GroupName       = $Group.displayName
        GroupId         = $Group.id
        Mail            = $Group.mail
        MailNickname    = $Group.mailNickname
        GroupTypes      = ($Group.groupTypes -join ";")
        DeletedDateTime = $DeletedAt.ToString("o")
        Status          = $Status
        Message         = $Message
    })
}

# ---------------- EXPORT REPORT ----------------

try {
    $Results | Export-Csv `
        -Path $ReportPath `
        -NoTypeInformation `
        -Encoding UTF8

    Write-Host ""
    Write-Host "Report saved: $ReportPath" -ForegroundColor Green
}
catch {
    Write-Host "REPORT ERROR: $($_.Exception.Message)" `
        -ForegroundColor Red
}

# ---------------- SUMMARY ----------------

$RestoredCount = @($Results | Where-Object Status -eq "Restored").Count
$FailedCount = @($Results | Where-Object Status -eq "Failed").Count
$PreviewCount = @($Results | Where-Object Status -eq "PreviewOnly").Count

Write-Host ""
Write-Host "=============== SUMMARY ==============="
Write-Host "Groups found:      $($Targets.Count)"
Write-Host "Restored:          $RestoredCount"
Write-Host "Failed:            $FailedCount"
Write-Host "Preview only:      $PreviewCount"
Write-Host "Report:            $ReportPath"
Write-Host "======================================="
