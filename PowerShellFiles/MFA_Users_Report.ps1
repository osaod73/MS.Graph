
# ==========================================================
# تقرير حالة المصادقة الثنائية MFA لمستخدمين من CSV
# Microsoft 365 / Microsoft Entra ID
# ==========================================================

$ErrorActionPreference = "Stop"

$InputFile = "C:\SchoolUsers\Users.csv"
$OutputFile = "C:\SchoolUsers\MFA_Users_Report.csv"

try {
    if (-not (Test-Path $InputFile)) {
        throw "ملف Users.csv غير موجود بجانب السكربت."
    }

    # تثبيت وحدة التقارير إذا لم تكن مثبتة
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Reports)) {
        Install-Module Microsoft.Graph.Reports `
            -Scope CurrentUser -Force
    }

    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Reports

    # الاتصال باستخدام مصادقة رمز الجهاز
    Connect-MgGraph -Scopes "AuditLog.Read.All" `
        -UseDeviceAuthentication

    # قراءة قائمة المستخدمين
    $InputUsers = @(Import-Csv -Path $InputFile)

    if ($InputUsers.Count -eq 0 -or
        -not ($InputUsers[0].PSObject.Properties.Name -contains "UserPrincipalName")) {
        throw "يجب أن يحتوي CSV على عمود UserPrincipalName."
    }

    $UPNs = @(
        $InputUsers |
        ForEach-Object {
            if (-not [string]::IsNullOrWhiteSpace($_.UserPrincipalName)) {
                $_.UserPrincipalName.Trim()
            }
        } |
        Sort-Object -Unique
    )

    Write-Host "جاري جلب تقرير تسجيل وسائل المصادقة..." `
        -ForegroundColor Cyan

    # جلب تقرير التسجيل من Microsoft Graph
    $RegistrationDetails = @(
        Get-MgReportAuthenticationMethodUserRegistrationDetail -All
    )

    # فهرسة النتائج حسب UPN
    $RegistrationIndex = @{}

    foreach ($Detail in $RegistrationDetails) {
        if (-not [string]::IsNullOrWhiteSpace($Detail.UserPrincipalName)) {
            $RegistrationIndex[$Detail.UserPrincipalName.ToLowerInvariant()] = $Detail
        }
    }

    # إنشاء التقرير للقائمة المحددة فقط
    $Report = foreach ($UPN in $UPNs) {
        $Key = $UPN.ToLowerInvariant()
        $Detail = $RegistrationIndex[$Key]

        if ($null -eq $Detail) {
            [PSCustomObject]@{
                UserPrincipalName = $UPN
                DisplayName       = ""
                MFARegistered     = "غير متاح"
                MFACapable        = "غير متاح"
                MethodsRegistered = ""
                Status            = "لم توجد بيانات في التقرير"
                LastUpdated       = ""
            }
            continue
        }

        $Methods = @($Detail.MethodsRegistered) -join "; "

        [PSCustomObject]@{
            UserPrincipalName = $Detail.UserPrincipalName
            DisplayName       = $Detail.UserDisplayName
            MFARegistered     = if ($Detail.IsMfaRegistered) {
                "نعم"
            } else {
                "لا"
            }
            MFACapable        = if ($Detail.IsMfaCapable) {
                "نعم"
            } else {
                "لا"
            }
            MethodsRegistered = $Methods
            Status            = if ($Detail.IsMfaCapable) {
                "مسجل وجاهز لاستخدام MFA"
            } elseif ($Detail.IsMfaRegistered) {
                "مسجل ولكن ليس جاهزاً حسب السياسة"
            } else {
                "لم يسجل MFA"
            }
            LastUpdated       = $Detail.LastUpdatedDateTime
        }
    }

    # حفظ التقرير وعرضه
    $Report |
        Export-Csv -Path $OutputFile `
            -NoTypeInformation -Encoding UTF8

    $Report |
        Format-Table UserPrincipalName, MFARegistered,
            MFACapable, Status -AutoSize

    Write-Host "`nتم حفظ التقرير في:" -ForegroundColor Green
    Write-Host $OutputFile -ForegroundColor Yellow
}
catch {
    Write-Host "حدث خطأ: $($_.Exception.Message)" `
        -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
