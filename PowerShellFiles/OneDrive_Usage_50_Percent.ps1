
# ==========================================================
# OneDrive Storage Usage Report - Users Using 50% or More
# Microsoft 365 / Microsoft Graph
# ==========================================================

$ErrorActionPreference = "Stop"

# نسبة الاستهلاك المطلوبة
$Threshold = 50

# مسار حفظ التقرير
$ReportPath = Join-Path $env:USERPROFILE "Desktop\OneDrive_Usage_50_Percent.csv"

try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

    # تثبيت الوحدة عند عدم وجودها
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        Install-Module Microsoft.Graph.Authentication `
            -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Authentication

    # اتصال باستخدام تسجيل الدخول عبر المتصفح
    Connect-MgGraph -Scopes "Reports.Read.All" -NoWelcome `
        -UseDeviceCode -ErrorAction Stop

    Write-Host "Downloading OneDrive usage report..." -ForegroundColor Cyan

    # تنزيل تقرير استخدام OneDrive
    $TempFile = Join-Path $env:TEMP "OneDriveUsageDetail.csv"

    Get-MgReportOneDriveUsageAccountDetail `
        -Period D7 `
        -OutFile $TempFile `
        -ErrorAction Stop

    # قراءة التقرير
    $Data = Import-Csv -Path $TempFile

    $Results = foreach ($User in $Data) {

        $UsedBytes = 0.0
        $AllocatedBytes = 0.0

        if (-not [double]::TryParse(
            [string]$User.'Storage Used (Byte)',
            [ref]$UsedBytes
        )) {
            continue
        }

        if (-not [double]::TryParse(
            [string]$User.'Storage Allocated (Byte)',
            [ref]$AllocatedBytes
        )) {
            continue
        }

        # تجاهل الحسابات التي لا تتوفر لها سعة مخصصة
        if ($AllocatedBytes -le 0) {
            continue
        }

        $Percentage = ($UsedBytes / $AllocatedBytes) * 100

        if ($Percentage -ge $Threshold) {

            [PSCustomObject]@{
                DisplayName       = $User.'Owner Display Name'
                UserPrincipalName = $User.'Owner Principal Name'
                StorageUsedGB     = [math]::Round(
                    $UsedBytes / 1GB, 2
                )
                StorageAllocatedGB = [math]::Round(
                    $AllocatedBytes / 1GB, 2
                )
                UsagePercentage   = [math]::Round(
                    $Percentage, 2
                )
                LastActivityDate  = $User.'Last Activity Date'
                SiteURL           = $User.'Site URL'
            }
        }
    }

    # ترتيب النتائج من الأعلى استهلاكاً
    $Results = @($Results | Sort-Object UsagePercentage -Descending)

    if ($Results.Count -gt 0) {

        $Results | Export-Csv -Path $ReportPath `
            -NoTypeInformation -Encoding UTF8

        $Results | Format-Table DisplayName,
            UserPrincipalName, StorageUsedGB,
            StorageAllocatedGB, UsagePercentage -AutoSize

        Write-Host ""
        Write-Host "Users found: $($Results.Count)" -ForegroundColor Green
        Write-Host "Report saved to: $ReportPath" -ForegroundColor Green
    }
    else {
        Write-Host "No users found at or above $Threshold% usage." `
            -ForegroundColor Yellow
    }

}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
