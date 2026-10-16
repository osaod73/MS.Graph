
# تقرير الوحدات الإدارية المحذوفة من Microsoft Entra ID
# Deleted Administrative Units Report
# https://login.microsoft.com/device

# تثبيت الوحدة عند الحاجة
if (-not (Get-Module -ListAvailable Microsoft.Graph.Identity.DirectoryManagement)) {
    Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser
}

Import-Module Microsoft.Graph.Identity.DirectoryManagement

# تسجيل الدخول بالصلاحيات المطلوبة
Connect-MgGraph -Scopes "Directory.Read.All" -UseDeviceCode

try {
    Write-Host "جاري البحث عن الوحدات الإدارية المحذوفة..." -ForegroundColor Cyan

    # استرجاع العناصر المحذوفة مؤقتاً
    $DeletedItems = Get-MgDirectoryDeletedItemAsAdministrativeUnit -All `
        -ErrorAction Stop

    # تجهيز التقرير
    $Report = foreach ($AU in $DeletedItems) {
        [PSCustomObject]@{
            DisplayName = $AU.DisplayName
            Id          = $AU.Id
            Description = $AU.Description
            Visibility  = $AU.Visibility
            DeletedDate = $AU.DeletedDateTime
            ReportDate  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
    }

    # إنشاء مجلد التقارير
    $ReportFolder = "$env:USERPROFILE\Documents\AUReports"

    if (-not (Test-Path $ReportFolder)) {
        New-Item -Path $ReportFolder -ItemType Directory -Force |
            Out-Null
    }

    # حفظ التقرير
    $ReportPath = Join-Path $ReportFolder (
        "Deleted_AU_Report_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
    )

    if ($Report) {
        $Report | Export-Csv -Path $ReportPath `
            -NoTypeInformation -Encoding UTF8

        Write-Host "تم إنشاء التقرير بنجاح." -ForegroundColor Green
        Write-Host "عدد الوحدات المحذوفة: $($Report.Count)"
        Write-Host "مسار التقرير: $ReportPath"
    }
    else {
        Write-Host "لم يتم العثور على وحدات إدارية محذوفة." `
            -ForegroundColor Yellow
    }
}
catch {
    Write-Host "حدث خطأ أثناء إنشاء التقرير:" -ForegroundColor Red
    Write-Host $_.Exception.Message
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
