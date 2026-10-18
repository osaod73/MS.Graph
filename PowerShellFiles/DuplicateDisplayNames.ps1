
# ==========================================================
# Report: Duplicate DisplayName in Microsoft 365
# ==========================================================

$ErrorActionPreference = "Stop"

# مسار حفظ التقرير
$ReportPath = Join-Path ([Environment]::GetFolderPath("Desktop")) "DuplicateDisplayNames.csv"

try {
    # تثبيت وحدة Microsoft Graph إذا لم تكن مثبتة
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force -AllowClobber
    }

    # تحميل الوحدة
    Import-Module Microsoft.Graph.Users -ErrorAction Stop

    # الاتصال بمايكروسوفت 365
   Connect-MgGraph -Scopes "User.Read.All" -UseDeviceAuthentication -NoWelcome

    Write-Host "جاري جلب جميع المستخدمين..." -ForegroundColor Cyan

    # جلب جميع المستخدمين مع الخصائص المطلوبة
    $Users = Get-MgUser -All -Property DisplayName, UserPrincipalName |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.DisplayName)
        }

    # تحديد أسماء العرض المكررة
    $DuplicateNames = $Users |
        Group-Object {
            $_.DisplayName.Trim()
        } |
        Where-Object { $_.Count -gt 1 }

    # إنشاء التقرير
    $Report = foreach ($Group in $DuplicateNames) {
        foreach ($User in $Group.Group) {
            [PSCustomObject]@{
                DisplayName       = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                DuplicateCount    = $Group.Count
            }
        }
    }

    if ($Report) {
        $Report |
            Sort-Object DisplayName, UserPrincipalName |
            Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

        Write-Host "`nتم العثور على أسماء مكررة." -ForegroundColor Green
        $Report | Format-Table DisplayName, UserPrincipalName, DuplicateCount -AutoSize
        Write-Host "`nتم حفظ التقرير في:" -ForegroundColor Yellow
        Write-Host $ReportPath
    }
    else {
        Write-Host "لم يتم العثور على أسماء عرض مكررة." -ForegroundColor Green
    }
}
catch {
    Write-Host "حدث خطأ: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
