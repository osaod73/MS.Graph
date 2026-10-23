# ============================================================
# Exclude CSV users from a specific MFA Conditional Access policy
# Microsoft Entra ID / Microsoft Graph
# ============================================================

$ErrorActionPreference = "Stop"

# Configuration
$CsvPath = "C:\SchoolUsers\MFAUsers.csv"
$OutputFolder = "C:\SchoolUsers"
$ReportPath = Join-Path $OutputFolder "Disable-MFA-Results.csv"

$PolicyName = "CA-Require-MFA-Selected-Users"

$Connected = $false

try {
    # --------------------------------------------------------
    # 1. Validate CSV and prepare report folder
    # --------------------------------------------------------

    if (-not (Test-Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force |
            Out-Null
    }

    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }

    $CsvUsers = @(
        Import-Csv -Path $CsvPath |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace(
                $_.UserPrincipalName
            )
        }
    )

    if ($CsvUsers.Count -eq 0) {
        throw "No valid UserPrincipalName entries found in CSV."
    }

    # --------------------------------------------------------
    # 2. Import installed modules only
    # --------------------------------------------------------

    $RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Identity.SignIns"
    )

    foreach ($ModuleName in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
            throw "Module '$ModuleName' is missing. Install it separately."
        }

        Import-Module $ModuleName -ErrorAction Stop
    }

    # --------------------------------------------------------
    # 3. Connect to Microsoft Graph
    # --------------------------------------------------------

    Connect-MgGraph -Scopes @(
        "User.Read.All",
        "Policy.Read.All",
        "Policy.ReadWrite.ConditionalAccess"
    ) -UseDeviceCode -NoWelcome

    $Connected = $true

    # --------------------------------------------------------
    # 4. Find the exact Conditional Access policy
    # --------------------------------------------------------

    $Policies = @(
        Get-MgIdentityConditionalAccessPolicy -All |
        Where-Object { $_.DisplayName -eq $PolicyName }
    )

    if ($Policies.Count -ne 1) {
        throw "Expected exactly one policy named '$PolicyName'."
    }

    $Policy = $Policies[0]

    if ($Policy.State -eq "disabled") {
        throw "The policy is already disabled. No changes made."
    }

    if (@($Policy.GrantControls.BuiltInControls) -notcontains "mfa") {
        throw "The policy does not explicitly require MFA."
    }

    Write-Host "Policy found: $($Policy.DisplayName)" `
        -ForegroundColor Cyan
    Write-Host "Policy ID: $($Policy.Id)"
    Write-Host "Current state: $($Policy.State)"

    # --------------------------------------------------------
    # 5. Resolve users before making any changes
    # --------------------------------------------------------

    $ResolvedUsers = @()
    $Results = @(
        foreach ($Entry in $CsvUsers) {
            $UPN = $Entry.UserPrincipalName.Trim()

            try {
                $DirectoryUser = Get-MgUser `
                    -UserId $UPN `
                    -Property Id,DisplayName,UserPrincipalName

                $ResolvedUsers += $DirectoryUser

                [PSCustomObject]@{
                    UserPrincipalName = $UPN
                    DisplayName       = $DirectoryUser.DisplayName
                    Status            = "Validated"
                    Details           = ""
                }
            }
            catch {
                [PSCustomObject]@{
                    UserPrincipalName = $UPN
                    DisplayName       = ""
                    Status            = "Failed"
                    Details           = $_.Exception.Message
                }
            }
        }
    )

    $Results | Export-Csv -Path $ReportPath `
        -NoTypeInformation -Encoding UTF8

    $FailedUsers = @(
        $Results | Where-Object { $_.Status -eq "Failed" }
    )

    if ($FailedUsers.Count -gt 0) {
        throw "One or more users could not be resolved. No policy changes made. Review $ReportPath."
    }

    # --------------------------------------------------------
    # 6. Retrieve policy directly as JSON
    # --------------------------------------------------------

    $PolicyUri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies/$($Policy.Id)"

    $CurrentPolicy = Invoke-MgGraphRequest `
        -Method GET `
        -Uri $PolicyUri `
        -ErrorAction Stop

    if ($CurrentPolicy.state -eq "disabled") {
        throw "The policy is disabled. No changes made."
    }

    if (@($CurrentPolicy.grantControls.builtInControls) -notcontains "mfa") {
        throw "The retrieved policy does not require MFA."
    }

    if (-not $CurrentPolicy.conditions.users) {
        throw "Policy user conditions could not be read."
    }

    # Keep all existing user exclusions
    $ExistingExclusions = @(
        $CurrentPolicy.conditions.users.excludeUsers |
        Where-Object { $_ }
    )

    $UserIdsToAdd = @(
        $ResolvedUsers |
        ForEach-Object { $_.Id }
    )

    $NewExclusions = @(
        $ExistingExclusions
        $UserIdsToAdd
    ) | Where-Object { $_ } | Select-Object -Unique

    # --------------------------------------------------------
    # 7. Confirmation
    # --------------------------------------------------------

    Write-Host ""
    Write-Host "Users to exclude from this policy: $($ResolvedUsers.Count)"
    Write-Host "Existing exclusions will be preserved."
    Write-Host "Other MFA policies may still apply."
    Write-Host ""

    $Confirmation = Read-Host `
        "Type YES to exclude the CSV users from this policy"

    if ($Confirmation -cne "YES") {
        Write-Host "Cancelled. No changes made."
        return
    }

    # --------------------------------------------------------
    # 8. Update only the policy conditions
    #    Preserve every existing condition returned by Graph.
    # --------------------------------------------------------

    $CurrentPolicy.conditions.users.excludeUsers = @($NewExclusions)

    $RequestBody = @{
        conditions = $CurrentPolicy.conditions
    } | ConvertTo-Json -Depth 50

    try {
        Invoke-MgGraphRequest `
            -Method PATCH `
            -Uri $PolicyUri `
            -Body $RequestBody `
            -ContentType "application/json" `
            -ErrorAction Stop | Out-Null
    }
    catch {
        Write-Host "Graph API rejected the policy update." `
            -ForegroundColor Red
        Write-Host $_.Exception.Message

        if ($_.ErrorDetails.Message) {
            Write-Host $_.ErrorDetails.Message
        }

        throw
    }

    # --------------------------------------------------------
    # 9. Verify the change
    # --------------------------------------------------------

    $VerifiedPolicy = Invoke-MgGraphRequest `
        -Method GET `
        -Uri $PolicyUri `
        -ErrorAction Stop

    $VerifiedExclusions = @(
        $VerifiedPolicy.conditions.users.excludeUsers
    )

    $FinalResults = foreach ($User in $ResolvedUsers) {
        $IsExcluded = $VerifiedExclusions -contains $User.Id

        [PSCustomObject]@{
            UserPrincipalName = $User.UserPrincipalName
            DisplayName       = $User.DisplayName
            Status            = if ($IsExcluded) {
                "Excluded from target policy"
            }
            else {
                "Verification failed"
            }
            Details = if ($IsExcluded) {
                "Other MFA policies may still apply."
            }
            else {
                "User ID was not found in verified exclusions."
            }
        }
    }

    $FinalResults | Export-Csv -Path $ReportPath `
        -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "========================================"
    Write-Host "OPERATION COMPLETED"
    Write-Host "========================================"
    Write-Host "Policy: $($VerifiedPolicy.displayName)"
    Write-Host "Policy state: $($VerifiedPolicy.state)"
    Write-Host "Report: $ReportPath"

    $FinalResults | Format-Table UserPrincipalName, Status -AutoSize

    if (@($FinalResults | Where-Object {
        $_.Status -eq "Verification failed"
    }).Count -gt 0) {
        Write-Host "WARNING: Some exclusions could not be verified." `
            -ForegroundColor Yellow
    }
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" `
        -ForegroundColor Red

    if ($_.ErrorDetails.Message) {
        Write-Host "Details: $($_.ErrorDetails.Message)" `
            -ForegroundColor Red
    }

    Write-Host "Review the report if it was created: $ReportPath"
}
finally {
    if ($Connected) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}