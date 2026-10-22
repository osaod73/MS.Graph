

# تقرير عدد مرات تغيير كلمات المرور بواسطة المسؤولين
# خلال آخر 30 يوماً - Microsoft 365 / Entra ID

$ErrorActionPreference = "Stop"

# ملف التقرير
$ReportPath = Join-Path $env:USERPROFILE "Desktop\PasswordChangesByAdmins_30Days.csv"

try {
    # تثبيت وحدة Graph عند الحاجة
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
    }

    Import-Module Microsoft.Graph.Authentication

    # الاتصال باستخدام Device Code لتجنب مشكلة WAM
    Connect-MgGraph -Scopes "AuditLog.Read.All" -UseDeviceCode -NoWelcome

    $StartDate = (Get-Date).ToUniversalTime().AddDays(-30)
    $StartDateString = $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")

    Write-Host "جاري جلب سجلات التدقيق..." -ForegroundColor Cyan

    # جلب سجلات التدقيق خلال آخر 30 يوماً
    $Uri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=activityDateTime ge $StartDateString&`$top=100"

    $AuditLogs = @()

    do {
        $Response = Invoke-MgGraphRequest -Method GET -Uri $Uri

        $AuditLogs += @($Response.value)

        $Uri = $Response.'@odata.nextLink'
    } while ($Uri)

    # عمليات كلمات المرور المحتملة
    $PasswordActivities = @(
        "Reset password",
        "Change password",
        "Admin registered security info",
        "Reset user password"
    )

    $Results = foreach ($Log in $AuditLogs) {

        # استبعاد العمليات الفاشلة
        if ($Log.result -ne "success") {
            continue
        }

        # التأكد من أن النشاط متعلق بكلمة المرور
        if ($Log.activityDisplayName -notin $PasswordActivities) {
            continue
        }

        # يجب أن يكون المنفذ مسؤولاً بشرياً وليس تطبيقاً
        $Admin = $Log.initiatedBy.user

        if (-not $Admin) {
            continue
        }

        # استخراج المستخدم المستهدف
        $Targets = @(
            $Log.targetResources |
            Where-Object { $_.type -eq "User" }
        )

        foreach ($Target in $Targets) {
            [PSCustomObject]@{
                AdminName       = $Admin.displayName
                AdminUPN        = $Admin.userPrincipalName
                TargetUser      = $Target.displayName
                TargetUPN       = $Target.userPrincipalName
                Activity        = $Log.activityDisplayName
                DateTimeUTC     = $Log.activityDateTime
                Result          = $Log.result
                AuditLogId      = $Log.id
            }
        }
    }

    # حفظ التفاصيل
    $Results = @($Results)

    $Results |
        Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

    # تلخيص عدد العمليات لكل مسؤول
    $SummaryPath = Join-Path $env:USERPROFILE "Desktop\PasswordChangesSummary_30Days.csv"

    $Summary = $Results |
        Group-Object AdminUPN |
        ForEach-Object {
            [PSCustomObject]@{
                AdminName = $_.Group[0].AdminName
                AdminUPN  = $_.Name
                TotalPasswordOperations = $_.Count
                UniqueUsers = @(
                    $_.Group.TargetUPN |
                    Where-Object { $_ } |
                    Sort-Object -Unique
                ).Count
            }
        } |
        Sort-Object TotalPasswordOperations -Descending

    $Summary |
        Export-Csv -Path $SummaryPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nاكتمل التقرير." -ForegroundColor Green
    Write-Host "تفاصيل العمليات: $ReportPath"
    Write-Host "ملخص المسؤولين: $SummaryPath"

    Write-Host "`nملخص عدد العمليات لكل مسؤول:" -ForegroundColor Cyan
    $Summary | Format-Table -AutoSize
}
catch {
    Write-Host "حدث خطأ: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
