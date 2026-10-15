
#Requires -Modules Microsoft.Graph.Users

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$CSVFilePath = "C:\SchoolUsers\UsersPasswords.csv",

    [switch]$ForceChangePasswordNextSignIn
)

$ErrorActionPreference = "Stop"

$requiredScopes = @(
    "User-PasswordProfile.ReadWrite.All",
    "User.Read.All"
)

# --------------------------------------------------
# 1. Connect to Microsoft Graph with required scopes
# --------------------------------------------------

try {
    $context = Get-MgContext

    $hasScopes = $false

    if ($context -and $context.Scopes) {
        $missingScopes = @(
            $requiredScopes | Where-Object {
                $_ -notin $context.Scopes
            }
        )

        $hasScopes = ($missingScopes.Count -eq 0)
    }

    if (-not $context -or -not $hasScopes) {
        if ($context) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        }

        Write-Host "Connecting to Microsoft Graph..." `
            -ForegroundColor Yellow

        Connect-MgGraph `
            -Scopes $requiredScopes `
            -UseDeviceCode `
            -ErrorAction Stop
    }

    $context = Get-MgContext

    if (-not $context) {
        throw "Microsoft Graph connection was not established."
    }

    $missingScopes = @(
        $requiredScopes | Where-Object {
            $_ -notin $context.Scopes
        }
    )

    if ($missingScopes.Count -gt 0) {
        throw "Missing Graph scopes: $($missingScopes -join ', '). Reconnect and grant consent."
    }

    Write-Host "Connected account: $($context.Account)" `
        -ForegroundColor Green

    Write-Host "Tenant ID: $($context.TenantId)" `
        -ForegroundColor Cyan
}
catch {
    throw "Graph authentication or permission check failed: $($_.Exception.Message)"
}

# --------------------------------------------------
# 2. Import and validate CSV
# --------------------------------------------------

if (-not (Test-Path $CSVFilePath -PathType Leaf)) {
    throw "CSV file not found: $CSVFilePath"
}

$users = @(Import-Csv -LiteralPath $CSVFilePath)

if ($users.Count -eq 0) {
    throw "The CSV file contains no user records."
}

$requiredColumns = @(
    "UserPrincipalName",
    "NewPassword"
)

$csvColumns = @($users[0].PSObject.Properties.Name)

foreach ($column in $requiredColumns) {
    if ($column -notin $csvColumns) {
        throw "CSV is missing required column: $column"
    }
}

$users = @(
    $users | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) -and
        -not [string]::IsNullOrEmpty($_.NewPassword)
    }
)

if ($users.Count -eq 0) {
    throw "No valid records found in the CSV."
}

$total = $users.Count
$index = 0

$results = [System.Collections.Generic.List[object]]::new()

Write-Host "`nUsers to process: $total" -ForegroundColor Cyan
Write-Host "Passwords will not be displayed or written to the log." `
    -ForegroundColor Cyan

# --------------------------------------------------
# 3. Update password profiles
# --------------------------------------------------

foreach ($user in $users) {
    $index++

    $upn = ([string]$user.UserPrincipalName).Trim()
    $newPassword = [string]$user.NewPassword

    Write-Progress `
        -Activity "Updating Microsoft 365 passwords" `
        -Status "[$index of $total] $upn" `
        -PercentComplete ([math]::Round(($index / $total) * 100))

    Write-Host "Processing: $upn" -ForegroundColor Cyan

    try {
        if ([string]::IsNullOrWhiteSpace($upn)) {
            throw "UserPrincipalName is empty."
        }

        $body = @{
            passwordProfile = @{
                password = $newPassword
                forceChangePasswordNextSignIn =
                    [bool]$ForceChangePasswordNextSignIn
            }
        }

        if ($PSCmdlet.ShouldProcess(
            $upn,
            "Update Microsoft Entra password"
        )) {
            Update-MgUser `
                -UserId $upn `
                -BodyParameter $body `
                -ErrorAction Stop

            Write-Host "[SUCCESS] $upn" -ForegroundColor Green

            $results.Add([PSCustomObject]@{
                UserPrincipalName = $upn
                Status            = "Success"
                Reason            = ""
                Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            })
        }
        else {
            $results.Add([PSCustomObject]@{
                UserPrincipalName = $upn
                Status            = "Skipped"
                Reason            = "Not executed (WhatIf or declined)"
                Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            })
        }
    }
    catch {
        $errorMessage = $_.Exception.Message

        Write-Host "[FAILED] $upn - $errorMessage" `
            -ForegroundColor Red

        $results.Add([PSCustomObject]@{
            UserPrincipalName = $upn
            Status            = "Failed"
            Reason            = $errorMessage
            Timestamp         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        })
    }
    finally {
        # Best-effort cleanup of local references
        $newPassword = $null
        $body = $null
    }
}

Write-Progress -Activity "Updating Microsoft 365 passwords" -Completed

# --------------------------------------------------
# 4. Export results without password values
# --------------------------------------------------

$logPath = Join-Path `
    (Split-Path -Parent $CSVFilePath) `
    ("PasswordResults_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss"))

$results | Export-Csv `
    -LiteralPath $logPath `
    -NoTypeInformation `
    -Encoding UTF8

$success = @($results | Where-Object Status -eq "Success").Count
$failed  = @($results | Where-Object Status -eq "Failed").Count
$skipped = @($results | Where-Object Status -eq "Skipped").Count

Write-Host "`n===== Summary =====" -ForegroundColor Cyan
Write-Host "Success : $success" -ForegroundColor Green
Write-Host "Failed  : $failed" -ForegroundColor Red
Write-Host "Skipped : $skipped" -ForegroundColor Yellow
Write-Host "Results : $logPath" -ForegroundColor Cyan
