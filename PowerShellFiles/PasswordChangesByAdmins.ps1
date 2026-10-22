
# Password Audit Report - Last 30 Days
# Microsoft Graph PowerShell
# Requires AuditLog.Read.All

$ErrorActionPreference = "Stop"

$Desktop = [Environment]::GetFolderPath("Desktop")
$DetailFile = Join-Path $Desktop "PasswordChangesByAdmins_30Days.csv"
$SummaryFile = Join-Path $Desktop "PasswordChangesSummary_30Days.csv"

$Connected = $false

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # اتصال مؤقت لتقليل مشكلات ذاكرة المصادقة
    Connect-MgGraph `
        -Scopes "AuditLog.Read.All" `
        -UseDeviceCode `
        -ContextScope Process `
        -NoWelcome `
        -ErrorAction Stop

    $Connected = $true

    $Context = Get-MgContext
    if (-not $Context -or -not $Context.Account) {
        throw "لم يتم إنشاء جلسة Microsoft Graph صالحة."
    }

    $StartDate = (Get-Date).ToUniversalTime().AddDays(-1)
    $FilterDate = [uri]::EscapeDataString(
        $StartDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
    )

    # حجم الصفحة صغير لتخفيف الضغط على Graph
    $Uri = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits?`$filter=activityDateTime%20ge%20$FilterDate&`$top=50"

    $Logs = [System.Collections.Generic.List[object]]::new()
    $Page = 0

    while ($Uri) {
        $Page++
        Write-Host "جلب الصفحة رقم $Page ..." -ForegroundColor Cyan

        $Response = $null
        $Success = $false

        for ($Attempt = 1; $Attempt -le 5; $Attempt++) {
            try {
                $Response = Invoke-MgGraphRequest `
                    -Method GET `
                    -Uri $Uri `
                    -ErrorAction Stop

                $Success = $true
                break
            }
            catch {
                if ($Attempt -eq 5) {
                    throw "فشل جلب الصفحة $Page بعد 5 محاولات. $($_.Exception.Message)"
                }

                $WaitSeconds = [Math]::Min(
                    60,
                    [Math]::Pow(2, $Attempt)
                )

                Write-Warning "فشل الطلب. المحاولة $Attempt؛ الانتظار $WaitSeconds ثانية."
                Start-Sleep -Seconds $WaitSeconds
            }
        }

        if (-not $Success) {
            throw "تعذر جلب سجلات التدقيق."
        }

        foreach ($Log in @($Response.value)) {
            $Logs.Add($Log)
        }

        # استخدام رابط الصفحة التالية كما أعاده Graph
        $Uri = $Response.'@odata.nextLink'
    }

    Write-Host "إجمالي سجلات التدقيق المسترجعة: $($Logs.Count)" `
        -ForegroundColor Green

    # ابدأ بعرض أسماء الأنشطة الفعلية المتعلقة بكلمات المرور
    $PasswordLogs = @(
        $Logs | Where-Object {
            $_.activityDisplayName -match '(?i)password'
        }
    )

    if ($PasswordLogs.Count -eq 0) {
        Write-Warning "لم تظهر أنشطة تحتوي كلمة password. راجع سجلات التدقيق وقائمة الأنشطة."
    }

    $Rows = foreach ($Log in $PasswordLogs) {

        if ($Log.result -ne "success") {
            continue
        }

        $Admin = $Log.initiatedBy.user

        # استبعاد العمليات التي نفذتها تطبيقات فقط
        if (-not $Admin) {
            continue
        }

        foreach ($Target in @($Log.targetResources)) {
            if ($Target.type -ne "User" -and $Target.Type -ne "User") {
                continue
            }

            [PSCustomObject]@{
                AdminName   = $Admin.displayName
                AdminUPN    = $Admin.userPrincipalName
                TargetUser  = $Target.displayName
                TargetUPN   = $Target.userPrincipalName
                Activity    = $Log.activityDisplayName
                DateTimeUTC = $Log.activityDateTime
                Result      = $Log.result
                AuditLogId  = $Log.id
            }
        }
    }

    $Rows = @($Rows)

    $Rows |
        Export-Csv -Path $DetailFile -NoTypeInformation -Encoding UTF8

    $Summary = @(
        $Rows |
        Group-Object AdminUPN |
        ForEach-Object {
            [PSCustomObject]@{
                AdminName = $_.Group[0].AdminName
                AdminUPN = $_.Name
                TotalOperations = $_.Count
                UniqueUsers = @(
                    $_.Group.TargetUPN |
                    Where-Object { $_ } |
                    Sort-Object -Unique
                ).Count
            }
        } |
        Sort-Object TotalOperations -Descending
    )

    $Summary |
        Export-Csv -Path $SummaryFile -NoTypeInformation -Encoding UTF8

    Write-Host "`nاكتمل التقرير." -ForegroundColor Green
    Write-Host "التفاصيل: $DetailFile"
    Write-Host "الملخص: $SummaryFile"

    $Summary | Format-Table -AutoSize
}
catch {
    Write-Host "`nحدث خطأ: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "إذا استمر الخطأ، افحص الاتصال والصلاحيات وسجلات التدقيق."
}
finally {
    if ($Connected) {
        try {
            Disconnect-MgGraph -ErrorAction Stop
        }
        catch {
            Write-Warning "تعذر إنهاء جلسة Graph بصورة طبيعية. يمكنك إغلاق نافذة PowerShell."
        }
    }
}
