
# ==========================================
# حذف مجموعة مستخدمين من Microsoft 365
# ==========================================

# https://login.microsoft.com/device

$CsvPath = "C:\SchoolUsers\users-to-delete.csv"

# تثبيت Microsoft Graph إذا لم يكن مثبتاً
if (-not (Get-Module Microsoft.Graph.Users -ListAvailable)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser
}

Import-Module Microsoft.Graph.Users

# الاتصال باستخدام Device Code لتجنب مشكلة WAM
Connect-MgGraph -Scopes "User.Read.All","User.ReadWrite.All" -UseDeviceCode

if (-not (Get-MgContext)) {
    throw "فشل الاتصال بـ Microsoft Graph"
}

# قراءة الحسابات
$Users = Import-Csv -Path $CsvPath

if (-not $Users -or
    -not ($Users[0].PSObject.Properties.Name -contains "UserPrincipalName")) {
    throw "ملف CSV فارغ أو لا يحتوي على عمود UserPrincipalName"
}

# عرض القائمة قبل الحذف
Write-Host "`nالحسابات المحددة للحذف:" -ForegroundColor Yellow
$Users | Format-Table UserPrincipalName

$ConfirmDelete = Read-Host "اكتب DELETE للتأكيد"

if ($ConfirmDelete -cne "DELETE") {
    Write-Host "تم إلغاء العملية."
    Disconnect-MgGraph
    return
}

# حذف الحسابات
foreach ($User in $Users) {
    $UPN = $User.UserPrincipalName.Trim()

    if ([string]::IsNullOrWhiteSpace($UPN)) {
        continue
    }

    try {
        # التحقق من وجود الحساب
        $Account = Get-MgUser -UserId $UPN -ErrorAction Stop

        # حذف الحساب
        Remove-MgUser -UserId $Account.Id -Confirm:$false -ErrorAction Stop

        Write-Host "تم حذف: $UPN" -ForegroundColor Green
    }
    catch {
        Write-Host "فشل حذف $UPN : $($_.Exception.Message)" `
            -ForegroundColor Red
    }
}

Disconnect-MgGraph
Write-Host "`nانتهت العملية."
