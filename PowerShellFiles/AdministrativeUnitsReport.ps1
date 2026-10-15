
# ==========================================================
# Export All Administrative Units from Microsoft Entra ID
# ==========================================================
# https://login.microsoft.com/device

# 1. Install required Graph modules if missing
$RequiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

foreach ($Module in $RequiredModules) {
    if (-not (Get-Module -ListAvailable -Name $Module)) {
        Install-Module $Module -Scope CurrentUser -Force
    }
    Import-Module $Module
}

# 2. Connect to Microsoft Graph
Disconnect-MgGraph -ErrorAction SilentlyContinue

Connect-MgGraph -Scopes @(
    "User.Read.All",
    "AdministrativeUnit.ReadWrite.All"
) -UseDeviceCode -NoWelcome

# Verify connection
Get-MgContext |
    Select-Object Account, TenantId, AuthType, Scopes

# مسار حفظ التقرير
$OutputFile = "C:\SchoolUsers\AdministrativeUnits.csv"

# إنشاء المجلد إذا لم يكن موجودًا
$OutputFolder = Split-Path $OutputFile -Parent

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

try {
    Write-Host "جاري جلب الوحدات الإدارية..." -ForegroundColor Cyan

    # جلب جميع الوحدات الإدارية مع دعم الصفحات
    $Units = Get-MgDirectoryAdministrativeUnit -All `
        -Property "id,displayName,description"

    # تجهيز التقرير
    $Report = foreach ($Unit in $Units) {
        [PSCustomObject]@{
            DisplayName = $Unit.DisplayName
            ObjectID    = $Unit.Id
            Description = $Unit.Description
        }
    }

    # تصدير التقرير إلى CSV
    $Report |
        Sort-Object DisplayName |
        Export-Csv -Path $OutputFile `
            -NoTypeInformation `
            -Encoding UTF8

    Write-Host ""
    Write-Host "تم استخراج التقرير بنجاح." -ForegroundColor Green
    Write-Host "عدد الوحدات الإدارية: $($Units.Count)"
    Write-Host "مسار التقرير: $OutputFile" -ForegroundColor Yellow

    # عرض النتائج على الشاشة
    $Report | Format-Table DisplayName, ObjectID, Description -AutoSize
}
catch {
    Write-Host "حدث خطأ أثناء استخراج التقرير:" -ForegroundColor Red
    Write-Host $_.Exception.Message
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
