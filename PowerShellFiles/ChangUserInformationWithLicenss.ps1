# ================================================================
# Microsoft Graph - Bulk User Update + License Assignment
# ================================================================

# -----------------------------
# Configuration
# -----------------------------

$CsvPath  = "C:\SchoolUsers\changuserinfo.csv"
$LogFile  = "C:\SchoolUsers\UpdateUsersLog102.txt"

# License SKU ID
$LicenseSkuId = "94763226-9b3c-4e75-a931-5c89701abe66"

# -----------------------------
# Connect to Microsoft Graph
# -----------------------------
# 1. إغلاق جلسة Microsoft Graph الحالية
Disconnect-MgGraph -ErrorAction SilentlyContinue

# 2. تعطيل WAM لهذه الجلسة
$env:MSAL_USE_WAM = "false"

# 3. الاتصال باستخدام رمز الجهاز
Connect-MgGraph -Scopes "User.ReadWrite.All","Organization.Read.All" -UseDeviceCode

# 4. التحقق من نجاح الاتصال
Get-MgContext

# https://login.microsoft.com/device

Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan

Connect-MgGraph -Scopes "User.ReadWrite.All","Organization.Read.All"

# -----------------------------
# Verify connection
# -----------------------------

$Context = Get-MgContext

if (-not $Context) {
    Write-Host "ERROR: Microsoft Graph connection failed." -ForegroundColor Red
    exit
}

Write-Host "Connected successfully." -ForegroundColor Green
Write-Host "Account: $($Context.Account)" -ForegroundColor Gray
Write-Host ""

# -----------------------------
# Check License SKU
# -----------------------------

Write-Host "Checking license SKU..." -ForegroundColor Cyan

try {

    $Sku = Get-MgSubscribedSku -All |
        Where-Object {
            $_.SkuId.ToString().ToLower() -eq $LicenseSkuId.ToLower()
        }

    if (-not $Sku) {
        Write-Host ""
        Write-Host "ERROR: The specified License SKU was not found in the tenant." -ForegroundColor Red
        Write-Host "SKU ID: $LicenseSkuId" -ForegroundColor Yellow
        exit
    }

    Write-Host "License found:" -ForegroundColor Green
    Write-Host "SKU Part Number : $($Sku.SkuPartNumber)" -ForegroundColor Gray
    Write-Host "Display Name    : $($Sku.SkuPartNumber)" -ForegroundColor Gray
    Write-Host "Available       : $($Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits)" -ForegroundColor Gray
    Write-Host ""

}
catch {

    Write-Host "ERROR: Unable to retrieve license information." -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit
}

# -----------------------------
# Check CSV
# -----------------------------

if (-not (Test-Path $CsvPath)) {

    Write-Host "ERROR: CSV file not found:" -ForegroundColor Red
    Write-Host $CsvPath -ForegroundColor Yellow
    exit
}

# -----------------------------
# Import CSV
# -----------------------------

$Users = Import-Csv -Path $CsvPath -Encoding UTF8

if (-not $Users) {

    Write-Host "ERROR: CSV file contains no users." -ForegroundColor Red
    exit
}

$TotalUsers = @($Users).Count
$Counter = 0

# -----------------------------
# Start Log
# -----------------------------

$StartTime = Get-Date

Add-Content -Path $LogFile -Value ""
Add-Content -Path $LogFile -Value "==============================================================="
Add-Content -Path $LogFile -Value "Bulk User Update and License Assignment"
Add-Content -Path $LogFile -Value "Start Time : $StartTime"
Add-Content -Path $LogFile -Value "CSV File   : $CsvPath"
Add-Content -Path $LogFile -Value "License SKU: $LicenseSkuId"
Add-Content -Path $LogFile -Value "==============================================================="

Write-Host ""
Write-Host "Starting Update Process..." -ForegroundColor Green
Write-Host "Total Users: $TotalUsers" -ForegroundColor Cyan
Write-Host ""

# ================================================================
# Process Users
# ================================================================

foreach ($user in $Users) {

    $Counter++

    $UserPrincipal = $null
    $UserObject = $null

    $UpdateStatus  = "Not Processed"
    $LicenseStatus = "Not Processed"

    Write-Host "---------------------------------------------------------------" -ForegroundColor DarkGray
    Write-Host "[$Counter / $TotalUsers]" -ForegroundColor Cyan

    # -----------------------------
    # Validate UPN
    # -----------------------------

    if ($null -eq $user.userPrincipalName) {

        Write-Host "ERROR: userPrincipalName is missing in CSV." -ForegroundColor Red

        Add-Content -Path $LogFile `
            -Value "$(Get-Date) | ERROR | Missing userPrincipalName in CSV row."

        continue
    }

    $UserPrincipal = $user.userPrincipalName.ToString().Trim()

    Write-Host "User: $UserPrincipal" -ForegroundColor White

    # ============================================================
    # Get User
    # ============================================================

    try {

        $UserObject = Get-MgUser `
            -UserId $UserPrincipal `
            -Property "Id,UserPrincipalName,UsageLocation,AssignedLicenses" `
            -ErrorAction Stop

    }
    catch {

        $Message = "$(Get-Date) | ERROR | $UserPrincipal | User not found or cannot be read | $($_.Exception.Message)"

        Write-Host "User not found / cannot be read." -ForegroundColor Red
        Add-Content -Path $LogFile -Value $Message

        continue
    }

    # ============================================================
    # Prepare User Information
    # ============================================================

    $JobTitle = ""

    if ($null -ne $user.jobTitle) {
        $JobTitle = $user.jobTitle.ToString().Trim()
    }

    # ============================================================
    # Update User Information
    # ============================================================

    try {

        $UserUpdateParams = @{
            UsageLocation     = "PS"
            City              = "مديرية التربية والتعليم /شمال الخليل"
            JobTitle          = "معلم/ة"
            Country           = "Palestinian Authority"
            PreferredLanguage = "ar-SA"
            PostalCode        = "970"
        }

        Update-MgUser `
            -UserId $UserPrincipal `
            -BodyParameter $UserUpdateParams `
            -ErrorAction Stop

        $UpdateStatus = "Success"

        Write-Host "User information: SUCCESS" -ForegroundColor Green

        Add-Content -Path $LogFile `
            -Value "$(Get-Date) | UPDATE SUCCESS | $UserPrincipal"

    }
    catch {

        $UpdateStatus = "Failed"

        $UpdateError = $_.Exception.Message

        Write-Host "User information: FAILED" -ForegroundColor Red
        Write-Host $UpdateError -ForegroundColor Red

        Add-Content -Path $LogFile `
            -Value "$(Get-Date) | UPDATE FAILED | $UserPrincipal | $UpdateError"
    }

    # ============================================================
    # License Assignment
    # ============================================================

    try {

        # Refresh user information after update
        $CurrentUser = Get-MgUser `
            -UserId $UserPrincipal `
            -Property "Id,UserPrincipalName,UsageLocation,AssignedLicenses" `
            -ErrorAction Stop

        # --------------------------------------------------------
        # Check UsageLocation
        # --------------------------------------------------------

        if ([string]::IsNullOrWhiteSpace($CurrentUser.UsageLocation)) {

            throw "UsageLocation is empty. License cannot be assigned."
        }

        # --------------------------------------------------------
        # Check if license already exists
        # --------------------------------------------------------

        $LicenseAlreadyAssigned = $false

        if ($CurrentUser.AssignedLicenses) {

            foreach ($AssignedLicense in $CurrentUser.AssignedLicenses) {

                if ($AssignedLicense.SkuId.ToString().ToLower() -eq $LicenseSkuId.ToLower()) {

                    $LicenseAlreadyAssigned = $true
                    break
                }
            }
        }

        # --------------------------------------------------------
        # License already assigned
        # --------------------------------------------------------

        if ($LicenseAlreadyAssigned) {

            $LicenseStatus = "Already Assigned"

            Write-Host "License: ALREADY ASSIGNED" -ForegroundColor Yellow

            Add-Content -Path $LogFile `
                -Value "$(Get-Date) | LICENSE ALREADY ASSIGNED | $UserPrincipal"

        }
        else {

            # ----------------------------------------------------
            # Check available licenses
            # ----------------------------------------------------

            $AvailableLicenses =
                $Sku.PrepaidUnits.Enabled - $Sku.ConsumedUnits

            if ($AvailableLicenses -le 0) {

                throw "No available licenses for SKU $LicenseSkuId."
            }

            # ----------------------------------------------------
            # Assign License
            # ----------------------------------------------------

            $LicenseParams = @{
                AddLicenses = @(
                    @{
                        SkuId = $LicenseSkuId
                    }
                )
                RemoveLicenses = @()
            }

            # Set-MgUserLicense does not print the Graph response
            $null = Set-MgUserLicense `
                -UserId $UserPrincipal `
                -BodyParameter $LicenseParams `
                -ErrorAction Stop

            $LicenseStatus = "Assigned"

            Write-Host "License: ASSIGNED SUCCESSFULLY" -ForegroundColor Green

            Add-Content -Path $LogFile `
                -Value "$(Get-Date) | LICENSE ASSIGNED | $UserPrincipal"
        }

    }
    catch {

        $LicenseStatus = "Failed"

        $LicenseError = $_.Exception.Message

        Write-Host "License: FAILED" -ForegroundColor Red
        Write-Host $LicenseError -ForegroundColor Red

        Add-Content -Path $LogFile `
            -Value "$(Get-Date) | LICENSE FAILED | $UserPrincipal | $LicenseError"
    }

    # ============================================================
    # Summary
    # ============================================================

    Write-Host ""
    Write-Host "Result:" -ForegroundColor Cyan
    Write-Host "  User Update : $UpdateStatus"
    Write-Host "  License     : $LicenseStatus"
    Write-Host ""

}

# ================================================================
# Finish
# ================================================================

$EndTime = Get-Date
$Duration = $EndTime - $StartTime

Add-Content -Path $LogFile -Value ""
Add-Content -Path $LogFile -Value "==============================================================="
Add-Content -Path $LogFile -Value "Process Completed"
Add-Content -Path $LogFile -Value "End Time : $EndTime"
Add-Content -Path $LogFile -Value "Duration : $Duration"
Add-Content -Path $LogFile -Value "Total Users : $TotalUsers"
Add-Content -Path $LogFile -Value "==============================================================="

Write-Host ""
Write-Host "===============================================================" -ForegroundColor Cyan
Write-Host "Finished Updating Users" -ForegroundColor Green
Write-Host "Total Users: $TotalUsers" -ForegroundColor Cyan
Write-Host "Duration   : $Duration" -ForegroundColor Cyan
Write-Host "Log File   : $LogFile" -ForegroundColor Cyan
Write-Host "===============================================================" -ForegroundColor Cyan