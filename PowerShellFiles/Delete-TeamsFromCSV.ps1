
# ==========================================
# Delete Microsoft Teams using Team IDs from CSV
# ==========================================
# https://login.microsoft.com/device

$InputFile  = "C:\SchoolUsers\teams-to-delete.csv"
$ReportFile = "C:\SchoolUsers\teams-to-delete_Report.csv"

# التأكد من وجود ملف الإدخال
if (-not (Test-Path $InputFile)) {
    Write-Host "ERROR: Teams.csv file not found: $InputFile" -ForegroundColor Red
    exit
}

# تحميل الوحدات المطلوبة
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Teams
Import-Module Microsoft.Graph.Groups

# تسجيل الدخول
Connect-MgGraph -Scopes "Team.ReadBasic.All", "Group.ReadWrite.All" -UseDeviceCode

# قراءة ملف CSV
$Teams = Import-Csv -Path $InputFile

if (-not ($Teams | Get-Member -Name TeamId -MemberType NoteProperty)) {
    Write-Host "ERROR: CSV must contain a TeamId column." -ForegroundColor Red
    Disconnect-MgGraph
    exit
}

$Results = foreach ($Item in $Teams) {

    $TeamId = "$($Item.TeamId)".Trim()

    if ([string]::IsNullOrWhiteSpace($TeamId)) {
        [PSCustomObject]@{
            TeamId = $TeamId
            TeamName = ""
            Status = "Failed"
            Message = "Empty TeamId"
            Date = Get-Date
        }
        continue
    }

    try {
        # التحقق من وجود الفريق
        $Team = Get-MgTeam -TeamId $TeamId -ErrorAction Stop

        Write-Host "`nTeam: $($Team.DisplayName)" -ForegroundColor Cyan
        Write-Host "ID: $TeamId"

        # تأكيد مستقل لكل فريق
        $Answer = Read-Host "Delete this team? Type YES to confirm"

        if ($Answer -cne "YES") {
            [PSCustomObject]@{
                TeamId = $TeamId
                TeamName = $Team.DisplayName
                Status = "Skipped"
                Message = "Deletion cancelled by user"
                Date = Get-Date
            }
            continue
        }

        # حذف مجموعة Microsoft 365 المرتبطة بالفريق
        Remove-MgGroup -GroupId $TeamId -Confirm:$false -ErrorAction Stop

        Write-Host "Deleted successfully: $TeamId" -ForegroundColor Green

        [PSCustomObject]@{
            TeamId = $TeamId
            TeamName = $Team.DisplayName
            Status = "Deleted"
            Message = "Team and associated Microsoft 365 group deletion requested"
            Date = Get-Date
        }
    }
    catch {
        Write-Host "FAILED: $TeamId - $($_.Exception.Message)" -ForegroundColor Red

        [PSCustomObject]@{
            TeamId = $TeamId
            TeamName = ""
            Status = "Failed"
            Message = $_.Exception.Message
            Date = Get-Date
        }
    }
}

# تصدير تقرير النتائج
$Results | Export-Csv -Path $ReportFile -NoTypeInformation -Encoding UTF8

Write-Host "`nProcess completed." -ForegroundColor Green
Write-Host "Report: $ReportFile" -ForegroundColor Yellow

Disconnect-MgGraph
