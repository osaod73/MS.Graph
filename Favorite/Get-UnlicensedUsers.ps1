
# تقرير المستخدمين الذين لا يمتلكون أي ترخيص في Microsoft 365
# https://login.microsoft.com/device

$ErrorActionPreference = "Stop"
$ReportPath = Join-Path $env:USERPROFILE "Desktop\UnlicensedUsers.csv"

try {
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
        Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
    }

    Import-Module Microsoft.Graph.Users

    Connect-MgGraph -Scopes "User.Read.All" -UseDeviceCode

    Write-Host "جاري جلب جميع المستخدمين..." -ForegroundColor Cyan

    $AllUsers = Get-MgUser -All -Property @(
        "Id",
        "DisplayName",
        "UserPrincipalName",
        "Mail",
        "Department",
        "JobTitle",
        "AccountEnabled",
        "UserType",
        "CreatedDateTime",
        "AssignedLicenses"
    )

    Write-Host "جاري تصفية المستخدمين بدون تراخيص..." -ForegroundColor Cyan

    $UnlicensedUsers = @(
        $AllUsers | Where-Object {
            $null -eq $_.AssignedLicenses -or
            $_.AssignedLicenses.Count -eq 0
        }
    )

    $Report = @(
        foreach ($User in $UnlicensedUsers) {
            [PSCustomObject]@{
                DisplayName       = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                Email             = $User.Mail
                Department        = $User.Department
                JobTitle          = $User.JobTitle
                AccountEnabled    = $User.AccountEnabled
                UserType          = $User.UserType
                CreatedDateTime   = $User.CreatedDateTime
                LicenseStatus     = "Unlicensed"
                UserId            = $User.Id
            }
        }
    )

    $Report | Export-Csv `
        -Path $ReportPath `
        -NoTypeInformation `
        -Encoding UTF8

    Write-Host ""
    Write-Host "تم إنشاء التقرير بنجاح." -ForegroundColor Green
    Write-Host "إجمالي المستخدمين: $($AllUsers.Count)"
    Write-Host "المستخدمون بدون ترخيص: $($Report.Count)"
    Write-Host "ملف التقرير: $ReportPath" -ForegroundColor Yellow

    $Report |
        Format-Table DisplayName, UserPrincipalName,
                     AccountEnabled, UserType -AutoSize
}
catch {
    Write-Host "حدث خطأ: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
