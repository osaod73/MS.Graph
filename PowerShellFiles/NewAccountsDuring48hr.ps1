# ============================================
# Microsoft 365 - New Accounts Report
# Accounts created during the last 48 hours
# ============================================

$ErrorActionPreference = "Stop"

# Report output path
$OutputFolder = "C:\SchoolUsers"
$OutputFile = Join-Path $OutputFolder (
    "NewAccounts_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
)

try {
    # Install Microsoft Graph if missing
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
        Write-Host "Installing Microsoft Graph Users module..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force -AllowClobber
    }

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Users

    # Sign in using device code
    Connect-MgGraph -Scopes "User.Read.All" -UseDeviceCode -NoWelcome

    # Calculate last 48 hours
    $StartDate = (Get-Date).ToUniversalTime().AddHours(-48)
    $EndDate   = (Get-Date).ToUniversalTime()

    Write-Host "Searching for accounts created during the last 48 hours..." -ForegroundColor Cyan
    Write-Host "From UTC: $StartDate"
    Write-Host "To UTC:   $EndDate"

    # Retrieve all users, then filter locally
    $Users = Get-MgUser -All -Property @(
        "id",
        "displayName",
        "userPrincipalName",
        "mail",
        "createdDateTime",
        "accountEnabled",
        "userType"
    )

    $NewUsers = $Users | Where-Object {
        $_.CreatedDateTime -and
        $_.CreatedDateTime.ToUniversalTime() -ge $StartDate -and
        $_.CreatedDateTime.ToUniversalTime() -le $EndDate
    } | Sort-Object CreatedDateTime -Descending

    # Create output folder
    if (-not (Test-Path -LiteralPath $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    # Prepare report
    $Report = $NewUsers | Select-Object `
        @{Name="DisplayName"; Expression={$_.DisplayName}},
        @{Name="UserPrincipalName"; Expression={$_.UserPrincipalName}},
        @{Name="Email"; Expression={
            if ($_.Mail) { $_.Mail } else { $_.UserPrincipalName }
        }},
        @{Name="CreatedDateTimeUTC"; Expression={
            $_.CreatedDateTime.ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss")
        }},
        @{Name="AccountEnabled"; Expression={$_.AccountEnabled}},
        @{Name="UserType"; Expression={$_.UserType}},
        @{Name="UserId"; Expression={$_.Id}}

    # Export CSV with UTF-8 encoding
    $Report | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Report completed successfully." -ForegroundColor Green
    Write-Host "New accounts found: $($Report.Count)" -ForegroundColor Cyan
    Write-Host "Report path: $OutputFile" -ForegroundColor Green

    # Display results
    $Report | Format-Table DisplayName, UserPrincipalName, CreatedDateTimeUTC, AccountEnabled -AutoSize

    # Open report automatically
    Invoke-Item -LiteralPath $OutputFile
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}