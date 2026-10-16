
# Microsoft 365 License Report
# Input: Users.csv
# Output: Microsoft365_License_Report.csv
# https://login.microsoft.com/device

$ErrorActionPreference = "Stop"

# تحديد مجلد العمل
if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $BasePath = $PSScriptRoot
}
else {
    $BasePath = (Get-Location).Path
}

$InputFile = Join-Path $BasePath "Users.csv"
$OutputFile = Join-Path $BasePath (
    "Microsoft365_License_Report_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
)

Write-Host "Input file:  $InputFile"
Write-Host "Output file: $OutputFile"

$InputFile = "C:\SchoolUsers\Users.csv"
$OutputFile = "C:\SchoolUsers\User_License_Report.csv"

# تثبيت الوحدة المطلوبة عند الحاجة
if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Users

$Results = [System.Collections.Generic.List[object]]::new()

try {
    if (-not (Test-Path $InputFile)) {
        throw "Users.csv not found: $InputFile"
    }

    # الاتصال بـ Microsoft Graph
    Connect-MgGraph -Scopes "User.Read.All" -UseDeviceCode -NoWelcome

    $Users = Import-Csv -Path $InputFile

    if (-not ($Users | Get-Member -Name UserPrincipalName -MemberType NoteProperty)) {
        throw "CSV must contain a UserPrincipalName column."
    }

    $Total = @($Users).Count
    $Counter = 0

    foreach ($Entry in $Users) {
        $Counter++
        $UPN = ([string]$Entry.UserPrincipalName).Trim()

        if ([string]::IsNullOrWhiteSpace($UPN)) {
            continue
        }

        Write-Progress `
            -Activity "Checking Microsoft 365 licenses" `
            -Status "$Counter of $Total : $UPN" `
            -PercentComplete (($Counter / [math]::Max($Total, 1)) * 100)

        try {
            # البحث عن المستخدم
            $EscapedUPN = $UPN.Replace("'", "''")

            $FoundUsers = @(
                Get-MgUser `
                    -Filter "userPrincipalName eq '$EscapedUPN'" `
                    -Property Id, DisplayName, UserPrincipalName, AccountEnabled `
                    -ErrorAction Stop
            )

            if ($FoundUsers.Count -eq 0) {
                $Results.Add([PSCustomObject]@{
                    InputUPN          = $UPN
                    DisplayName       = ""
                    UserPrincipalName = ""
                    AccountEnabled    = ""
                    LicenseName       = ""
                    SkuId             = ""
                    ServicePlans      = ""
                    LicenseStatus     = "User Not Found"
                    CheckedAt         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                })
                continue
            }

            foreach ($User in $FoundUsers) {
                # جلب تفاصيل التراخيص
                $LicenseDetails = @(
                    Get-MgUserLicenseDetail -UserId $User.Id -ErrorAction Stop
                )

                if ($LicenseDetails.Count -eq 0) {
                    $Results.Add([PSCustomObject]@{
                        InputUPN          = $UPN
                        DisplayName       = $User.DisplayName
                        UserPrincipalName = $User.UserPrincipalName
                        AccountEnabled    = $User.AccountEnabled
                        LicenseName       = "No License"
                        SkuId             = ""
                        ServicePlans      = ""
                        LicenseStatus     = "Unlicensed"
                        CheckedAt         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    })
                }
                else {
                    foreach ($License in $LicenseDetails) {
                        $Plans = @(
                            $License.ServicePlans |
                            ForEach-Object { $_.ServicePlanName }
                        ) -join "; "

                        $Results.Add([PSCustomObject]@{
                            InputUPN          = $UPN
                            DisplayName       = $User.DisplayName
                            UserPrincipalName = $User.UserPrincipalName
                            AccountEnabled    = $User.AccountEnabled
                            LicenseName       = $License.SkuPartNumber
                            SkuId             = $License.SkuId
                            ServicePlans      = $Plans
                            LicenseStatus     = "Licensed"
                            CheckedAt         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                        })
                    }
                }
            }
        }
        catch {
            $Results.Add([PSCustomObject]@{
                InputUPN          = $UPN
                DisplayName       = ""
                UserPrincipalName = ""
                AccountEnabled    = ""
                LicenseName       = ""
                SkuId             = ""
                ServicePlans      = ""
                LicenseStatus     = "Check Failed: $($_.Exception.Message)"
                CheckedAt         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            })

            Write-Warning "Failed to check ${UPN}: $($_.Exception.Message)"
        }
    }

    Write-Progress -Activity "Checking Microsoft 365 licenses" -Completed

    # تصدير التقرير
    $Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

    Write-Host "`nReport created successfully." -ForegroundColor Green
    Write-Host "Report: $OutputFile"
    Write-Host "Total input rows: $Total"
    Write-Host "Licensed accounts: $(@($Results | Where-Object LicenseStatus -eq 'Licensed' | Select-Object UserPrincipalName -Unique).Count)"
    Write-Host "Unlicensed accounts: $(@($Results | Where-Object LicenseStatus -eq 'Unlicensed').Count)"
    Write-Host "Not found: $(@($Results | Where-Object LicenseStatus -eq 'User Not Found').Count)"
    Write-Host "Failed checks: $(@($Results | Where-Object { $_.LicenseStatus -like 'Check Failed:*' }).Count)"
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
