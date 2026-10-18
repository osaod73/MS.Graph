# ============================================================
# Microsoft 365 Failed Sign-in Report
# Last 24 Hours
# Includes Department and Job Title
# ============================================================

$ErrorActionPreference = "Stop"

# Report settings
$OutputFolder = "C:\M365Reports"
$TimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$OutputFile = Join-Path $OutputFolder "FailedSignIns_$TimeStamp.csv"

try {
    # --------------------------------------------------------
    # 1. Install required Microsoft Graph modules
    # --------------------------------------------------------

    $Modules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Reports",
        "Microsoft.Graph.Users"
    )

    foreach ($Module in $Modules) {

        if (-not (Get-Module -ListAvailable -Name $Module)) {

            Write-Host "Installing $Module ..." -ForegroundColor Yellow

            Install-Module $Module `
                -Scope CurrentUser `
                -Force `
                -AllowClobber
        }

        Import-Module $Module -ErrorAction Stop
    }

    # --------------------------------------------------------
    # 2. Connect to Microsoft Graph
    # --------------------------------------------------------

    Connect-MgGraph -Scopes @(
        "AuditLog.Read.All",
        "User.Read.All"
    ) -UseDeviceCode -NoWelcome

    # --------------------------------------------------------
    # 3. Calculate the last 24 hours
    # --------------------------------------------------------

    $EndDate = (Get-Date).ToUniversalTime()
    $StartDate = $EndDate.AddHours(-24)

    $StartText = $StartDate.ToString(
        "yyyy-MM-ddTHH:mm:ssZ",
        [Globalization.CultureInfo]::InvariantCulture
    )

    $EndText = $EndDate.ToString(
        "yyyy-MM-ddTHH:mm:ssZ",
        [Globalization.CultureInfo]::InvariantCulture
    )

    $Filter = "createdDateTime ge $StartText and createdDateTime le $EndText"

    Write-Host ""
    Write-Host "Searching failed sign-in attempts..." -ForegroundColor Cyan
    Write-Host "From UTC: $StartText"
    Write-Host "To UTC:   $EndText"

    # --------------------------------------------------------
    # 4. Retrieve sign-in logs
    # --------------------------------------------------------

    $SignIns = Get-MgAuditLogSignIn -Filter $Filter -All

    # 50126 = Invalid username or password
    # 50034 = User account does not exist

    $FailedSignIns = @(
        $SignIns |
            Where-Object {
                $_.Status.ErrorCode -in @(50126, 50034)
            } |
            Sort-Object CreatedDateTime -Descending
    )

    # --------------------------------------------------------
    # 5. Create output folder
    # --------------------------------------------------------

    if (-not (Test-Path -LiteralPath $OutputFolder)) {

        New-Item -Path $OutputFolder `
            -ItemType Directory `
            -Force | Out-Null
    }

    # --------------------------------------------------------
    # 6. Cache user details to reduce Graph requests
    # --------------------------------------------------------

    $UserCache = @{}

    # --------------------------------------------------------
    # 7. Build the report
    # --------------------------------------------------------

    $Report = foreach ($SignIn in $FailedSignIns) {

        $UserId = $SignIn.UserId
        $Department = ""
        $JobTitle = ""

        # Retrieve department and job title when UserId exists
        if (-not [string]::IsNullOrWhiteSpace($UserId)) {

            if (-not $UserCache.ContainsKey($UserId)) {

                try {

                    $UserInfo = Get-MgUser `
                        -UserId $UserId `
                        -Property @(
                            "id",
                            "displayName",
                            "userPrincipalName",
                            "department",
                            "jobTitle"
                        ) `
                        -ErrorAction Stop

                    $UserCache[$UserId] = $UserInfo
                }
                catch {

                    # User may be unknown, deleted, or inaccessible
                    $UserCache[$UserId] = $null
                }
            }

            $CachedUser = $UserCache[$UserId]

            if ($null -ne $CachedUser) {

                $Department = $CachedUser.Department
                $JobTitle = $CachedUser.JobTitle
            }
        }

        # Translate error codes
        $ErrorDescription = switch ($SignIn.Status.ErrorCode) {
            50126 { "Invalid username or password" }
            50034 { "User account does not exist" }
            default { $SignIn.Status.FailureReason }
        }

        # Create report row
        [PSCustomObject]@{

            DateTimeUTC = if ($SignIn.CreatedDateTime) {
                $SignIn.CreatedDateTime.ToUniversalTime().
                    ToString("yyyy-MM-dd HH:mm:ss")
            } else {
                ""
            }

            DisplayName = $SignIn.UserDisplayName

            UserPrincipalName = $SignIn.UserPrincipalName

            SignInIdentifier = $SignIn.SignInIdentifier

            Department = $Department

            JobTitle = $JobTitle

            IPAddress = $SignIn.IpAddress

            Country = $SignIn.Location.CountryOrRegion

            City = $SignIn.Location.City

            Application = $SignIn.AppDisplayName

            ClientApp = $SignIn.ClientAppUsed

            ErrorCode = $SignIn.Status.ErrorCode

            ErrorDescription = $ErrorDescription

            FailureReason = $SignIn.Status.FailureReason

            AdditionalDetails = $SignIn.Status.AdditionalDetails

            CorrelationId = $SignIn.CorrelationId
        }
    }

    # --------------------------------------------------------
    # 8. Export CSV report
    # --------------------------------------------------------

    $Report | Export-Csv `
        -Path $OutputFile `
        -NoTypeInformation `
        -Encoding UTF8

    # --------------------------------------------------------
    # 9. Display summary
    # --------------------------------------------------------

    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Report created successfully!" -ForegroundColor Green
    Write-Host "Failed attempts: $($Report.Count)" -ForegroundColor Yellow
    Write-Host "Report path: $OutputFile" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Green

    # Display results
    $Report |
        Format-Table DateTimeUTC,
            DisplayName,
            UserPrincipalName,
            Department,
            JobTitle,
            IPAddress,
            ErrorCode -AutoSize

    # Open the CSV report
    Invoke-Item -LiteralPath $OutputFile
}
catch {

    Write-Host ""
    Write-Host "An error occurred:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {

    Disconnect-MgGraph -ErrorAction SilentlyContinue
}