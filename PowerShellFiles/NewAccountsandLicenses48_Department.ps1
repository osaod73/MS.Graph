# Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
# ============================================
# Microsoft 365 - New Accounts and Licenses Report
# Last 48 hours
# ============================================

$ErrorActionPreference = "Stop"

$OutputFolder = "C:\SchoolUsers"
$OutputFile = Join-Path $OutputFolder (
    "NewAccounts_Licenses_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
)

try {
    # Install required modules if missing
    $RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Identity.DirectoryManagement"
    )

    foreach ($Module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            Write-Host "Installing $Module ..." -ForegroundColor Yellow
            Install-Module $Module -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module $Module -ErrorAction Stop
    }

    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes @(
        "User.Read.All",
        "LicenseAssignment.Read.All"
    ) -UseDeviceCode -NoWelcome

    # Time range: last 48 hours
    $StartDate = (Get-Date).ToUniversalTime().AddHours(-48)
    $EndDate = (Get-Date).ToUniversalTime()

    Write-Host "Retrieving users and license information..." -ForegroundColor Cyan

    # Get license catalog once
    $SkuCatalog = Get-MgSubscribedSku -All -ErrorAction Stop

    # Map SKU IDs to license product names
    $SkuMap = @{}
    foreach ($Sku in $SkuCatalog) {
        $SkuMap[$Sku.SkuId.ToString()] = $Sku.SkuPartNumber
    }

    # Retrieve users including assigned licenses
    $Users = Get-MgUser -All -Property @(
        "Id",
        "DisplayName",
        "UserPrincipalName",
        "Mail",
        "CreatedDateTime",
        "AccountEnabled",
        "UserType",
        "Department",
        "AssignedLicenses"
    )

    # Filter accounts created in the last 48 hours
    $NewUsers = @(
        $Users | Where-Object {
            $_.CreatedDateTime -and
            $_.CreatedDateTime.ToUniversalTime() -ge $StartDate -and
            $_.CreatedDateTime.ToUniversalTime() -le $EndDate
        } | Sort-Object CreatedDateTime -Descending
    )

    # Create output directory
    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force |
            Out-Null
    }

    # Build report with license names
    $Report = foreach ($User in $NewUsers) {

        $LicenseNames = @(
            foreach ($License in $User.AssignedLicenses) {
                $SkuId = $License.SkuId.ToString()

                if ($SkuMap.ContainsKey($SkuId)) {
                    $SkuMap[$SkuId]
                }
                else {
                    "Unknown SKU: $SkuId"
                }
            }
        ) | Sort-Object -Unique

        if ($LicenseNames.Count -eq 0) {
            $LicenseText = "Unlicensed"
        }
        else {
            $LicenseText = $LicenseNames -join ", "
        }

        [PSCustomObject]@{
            DisplayName       = $User.DisplayName
            UserPrincipalName = $User.UserPrincipalName
            Email             = if ($User.Mail) {
                                    $User.Mail
                                } else {
                                    $User.UserPrincipalName
                                }
            CreatedDateTimeUTC = $User.CreatedDateTime.
                ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
            AccountEnabled    = $User.AccountEnabled
            UserType           = $User.UserType
            Department         = if ($User.Department) { $User.Department } else { "Not specified" }
            License            = $LicenseText
            UserId             = $User.Id
        }
    }

    # Export report
    $Report | Export-Csv -Path $OutputFile `
        -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Report created successfully." -ForegroundColor Green
    Write-Host "Accounts found: $($Report.Count)" -ForegroundColor Cyan
    Write-Host "Report location: $OutputFile" -ForegroundColor Green

    # Display results
    $Report | Format-Table DisplayName, UserPrincipalName, Department, License,
        CreatedDateTimeUTC, AccountEnabled -AutoSize

    # Open CSV report
    Invoke-Item -LiteralPath $OutputFile
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
