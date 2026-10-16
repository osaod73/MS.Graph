
# Restore Deleted Administrative Unit (AU)
# Microsoft Entra ID / Microsoft Graph
# https://login.microsoft.com/device

$ErrorActionPreference = "Stop"

# تثبيت الوحدة المطلوبة إذا لم تكن مثبتة
if (-not (Get-Module -ListAvailable Microsoft.Graph.Identity.DirectoryManagement)) {
    Install-Module Microsoft.Graph.Identity.DirectoryManagement `
        -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Identity.DirectoryManagement

try {
    # تسجيل الدخول
    Connect-MgGraph -Scopes "AdministrativeUnit.ReadWrite.All" `
        -UseDeviceCode -NoWelcome

    Write-Host "`nجاري جلب الوحدات الإدارية المحذوفة..." `
        -ForegroundColor Cyan

    # جلب العناصر المحذوفة مؤقتاً
    $DeletedAUs = @(
        Get-MgDirectoryDeletedItemAsAdministrativeUnit -All
    )

    if ($DeletedAUs.Count -eq 0) {
        Write-Host "لا توجد وحدات إدارية محذوفة قابلة للاستعادة." `
            -ForegroundColor Yellow
        return
    }

    # عرض الوحدات المحذوفة
    $DeletedAUs |
        Select-Object DisplayName, Id, Description |
        Format-Table -AutoSize

    # إدخال معرف الوحدة المراد استعادتها
    $AUId = Read-Host "`nأدخل ID للوحدة الإدارية المراد استعادتها"

    $SelectedAU = $DeletedAUs |
        Where-Object { $_.Id -eq $AUId }

    if (-not $SelectedAU) {
        throw "لم يتم العثور على وحدة محذوفة بهذا المعرف."
    }

    Write-Host "`nالوحدة المحددة: $($SelectedAU.DisplayName)" `
        -ForegroundColor Cyan

    $ConfirmRestore = Read-Host "هل تريد استعادتها؟ اكتب YES للتأكيد"

    if ($ConfirmRestore -cne "YES") {
        Write-Host "تم إلغاء عملية الاستعادة." `
            -ForegroundColor Yellow
        return
    }

    # استعادة الوحدة باستخدام Microsoft Graph
    Restore-MgDirectoryDeletedItem `
        -DirectoryObjectId $SelectedAU.Id `
        -ErrorAction Stop

    Write-Host "`nتم إرسال طلب الاستعادة بنجاح." `
        -ForegroundColor Green

    # التحقق من وجود الوحدة بعد الاستعادة
    $RestoredAU = Get-MgDirectoryAdministrativeUnit `
        -AdministrativeUnitId $SelectedAU.Id `
        -ErrorAction Stop

    Write-Host "تم التحقق من استعادة الوحدة الإدارية:" `
        -ForegroundColor Green

    $RestoredAU |
        Select-Object DisplayName, Id, Description |
        Format-List
}
catch {
    Write-Host "`nفشلت العملية: $($_.Exception.Message)" `
        -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
