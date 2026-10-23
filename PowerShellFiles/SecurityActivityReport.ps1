# تقرير الأنشطة الأمنية المشبوهة في Microsoft 365
# احفظ الملف باسم SecurityActivityReport.ps1

$ErrorActionPreference = "Stop"

# الإعدادات
$DaysBack = 30
$OutputFolder = "C:\SchoolUsers\SecurityReport"

# إنشاء مجلد النتائج
if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$DateFrom = (Get-Date).ToUniversalTime().AddDays(-$DaysBack)
$DateTo = (Get-Date).ToUniversalTime()
$FromFilter = $DateFrom.ToString("yyyy-MM-ddTHH:mm:ssZ")
$ToFilter = $DateTo.ToString("yyyy-MM-ddTHH:mm:ssZ")

$ReportPath = Join-Path $OutputFolder "SecurityActivityReport.csv"
$EventsPath = Join-Path $OutputFolder "SecurityEventsDetails.csv"

try {
    # تثبيت وحدة Graph إذا لم تكن مثبتة
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Authentication

    # تسجيل الدخول بطريقة Device Code لتجنب مشكلة WAM
    Connect-MgGraph -Scopes "AuditLog.Read.All" -UseDeviceCode -NoWelcome

    # قراءة سجلات تسجيل الدخول عبر Microsoft Graph REST API
    $Uri = "https://graph.microsoft.com/v1.0/auditLogs/signIns?`$filter=createdDateTime ge $FromFilter and createdDateTime le $ToFilter&`$top=1000"

    $AllEvents = [System.Collections.Generic.List[object]]::new()

    do {
        Write-Host "جاري تحميل سجلات تسجيل الدخول..." -ForegroundColor Cyan

        $Response = Invoke-MgGraphRequest -Method GET -Uri $Uri

        foreach ($Event in $Response.value) {
            $AllEvents.Add($Event)
        }

        $Uri = $Response.'@odata.nextLink'
    } while ($Uri)

    # تحديد الأحداث المشبوهة
    $SuspiciousEvents = foreach ($Event in $AllEvents) {
        $Failure = (
            $null -ne $Event.status -and
            [int]$Event.status.errorCode -ne 0
        )

        $RiskLevel = [string]$Event.riskLevelAggregated
        $RiskDuringSignIn = [string]$Event.riskLevelDuringSignIn
        $RiskState = [string]$Event.riskState

        $HighRisk = $RiskLevel -in @("high", "medium") -or
                    $RiskDuringSignIn -in @("high", "medium") -or
                    $RiskState -in @("atRisk", "confirmedCompromised")

        if ($Failure -or $HighRisk) {
            [PSCustomObject]@{
                DateTimeUTC       = $Event.createdDateTime
                UserPrincipalName = $Event.userPrincipalName
                DisplayName       = $Event.userDisplayName
                AppName           = $Event.appDisplayName
                IPAddress         = $Event.ipAddress
                Country           = $Event.location.countryOrRegion
                City              = $Event.location.city
                ErrorCode         = $Event.status.errorCode
                FailureReason     = $Event.status.failureReason
                RiskLevel         = $RiskLevel
                RiskDuringSignIn  = $RiskDuringSignIn
                RiskState         = $RiskState
                RiskEventTypes    = (@($Event.riskEventTypes) -join "; ")
                IsInteractive     = $Event.isInteractive
            }
        }
    }

    # تصدير التفاصيل
    $SuspiciousEvents |
        Sort-Object DateTimeUTC -Descending |
        Export-Csv -Path $EventsPath -NoTypeInformation -Encoding UTF8

    # ترتيب المستخدمين حسب عدد الأحداث
    $Summary = $SuspiciousEvents |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) } |
        Group-Object UserPrincipalName |
        ForEach-Object {
            $UserEvents = @($_.Group)

            $FailedCount = @(
                $UserEvents | Where-Object {
                    [int]$_.ErrorCode -ne 0
                }
            ).Count

            $RiskCount = @(
                $UserEvents | Where-Object {
                    $_.RiskLevel -in @("high", "medium") -or
                    $_.RiskDuringSignIn -in @("high", "medium") -or
                    $_.RiskState -in @("atRisk", "confirmedCompromised")
                }
            ).Count

            [PSCustomObject]@{
                UserPrincipalName    = $_.Name
                DisplayName          = $UserEvents[0].DisplayName
                TotalSuspiciousEvents = $UserEvents.Count
                FailedSignIns        = $FailedCount
                RiskRelatedEvents    = $RiskCount
                UniqueIPAddresses    = @(
                    $UserEvents |
                    Where-Object { $_.IPAddress } |
                    Select-Object -ExpandProperty IPAddress -Unique
                ).Count
                LastEventUTC         = (
                    $UserEvents |
                    Sort-Object DateTimeUTC -Descending |
                    Select-Object -First 1
                ).DateTimeUTC
            }
        } |
        Sort-Object TotalSuspiciousEvents -Descending

    $Summary |
        Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nاكتمل التقرير بنجاح." -ForegroundColor Green
    Write-Host "تقرير المستخدمين: $ReportPath"
    Write-Host "تفاصيل الأحداث:   $EventsPath"

    Write-Host "`nأكثر 20 حسابًا حسب عدد الأحداث:" -ForegroundColor Yellow

    $Summary | Select-Object -First 20 |
        Format-Table UserPrincipalName, TotalSuspiciousEvents,
            FailedSignIns, RiskRelatedEvents, UniqueIPAddresses,
            LastEventUTC -AutoSize

}
catch {
    Write-Host "حدث خطأ: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "تحقق من الاتصال والصلاحيات والترخيص وتوفر الوحدة المطلوبة."
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}