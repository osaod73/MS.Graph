
# ============================================================
# Permanently delete all soft-deleted Entra ID user accounts
# WARNING: IRREVERSIBLE
# ============================================================

$ErrorActionPreference = "Stop"

$ReportPath = "C:\SchoolUsers\PermanentlyDeletedUsersReport.csv"
$ErrorLogPath = "C:\SchoolUsers\PermanentlyDeleteUsersErrors.txt"

$Results = [System.Collections.Generic.List[object]]::new()
$Errors = [System.Collections.Generic.List[string]]::new()

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # Workaround for WAM window-handle issues
    $env:MSAL_USE_WAM = "false"

    Connect-MgGraph `
        -Scopes "User.ReadWrite.All" `
        -UseDeviceCode `
        -NoWelcome `
        -ErrorAction Stop

    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green

    # Retrieve all deleted users, following pagination
    $uri = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,deletedDateTime&`$top=999"

    $DeletedUsers = [System.Collections.Generic.List[object]]::new()

    while (-not [string]::IsNullOrWhiteSpace($uri)) {
        $response = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $uri `
            -ErrorAction Stop

        foreach ($user in $response.value) {
            $DeletedUsers.Add($user)
        }

        $uri = $response.'@odata.nextLink'
    }

    if ($DeletedUsers.Count -eq 0) {
        Write-Host "No deleted user accounts were found." `
            -ForegroundColor Yellow
        return
    }

    Write-Host "`nDeleted user accounts found: $($DeletedUsers.Count)" `
        -ForegroundColor Yellow

    $DeletedUsers |
        Select-Object displayName, userPrincipalName, mail,
                      deletedDateTime, id |
        Format-Table -AutoSize

    Write-Host "`nWARNING: Permanent deletion cannot be undone." `
        -ForegroundColor Red

    $confirmation = Read-Host `
        "Type YES to permanently delete ALL listed user accounts"

    if ($confirmation -cne "YES") {
        Write-Host "Cancelled. No accounts were deleted." `
            -ForegroundColor Yellow
        return
    }

    # Permanently delete each user
    foreach ($user in $DeletedUsers) {
        try {
            $deleteUri = "https://graph.microsoft.com/v1.0/directory/deletedItems/$($user.id)"

            Invoke-MgGraphRequest `
                -Method DELETE `
                -Uri $deleteUri `
                -ErrorAction Stop

            $Results.Add([PSCustomObject]@{
                DisplayName       = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                Mail              = $user.mail
                UserId            = $user.id
                Status            = "Permanently deleted"
                Error             = ""
            })

            Write-Host "DELETED: $($user.userPrincipalName)" `
                -ForegroundColor Green
        }
        catch {
            $message = $_.Exception.Message

            if ($_.ErrorDetails.Message) {
                $message += " | " + $_.ErrorDetails.Message
            }

            $Errors.Add(
                "UserId=$($user.id); UPN=$($user.userPrincipalName); Error=$message"
            )

            $Results.Add([PSCustomObject]@{
                DisplayName       = $user.displayName
                UserPrincipalName = $user.userPrincipalName
                Mail              = $user.mail
                UserId            = $user.id
                Status            = "Failed"
                Error             = $message
            })

            Write-Host "FAILED: $($user.userPrincipalName)" `
                -ForegroundColor Red
            Write-Host $message
        }
    }
}
catch {
    $message = $_.Exception.Message

    if ($_.ErrorDetails.Message) {
        $message += " | " + $_.ErrorDetails.Message
    }

    $Errors.Add("SCRIPT ERROR: $message")
    Write-Host "SCRIPT ERROR: $message" -ForegroundColor Red
}
finally {
    if ($Results.Count -gt 0) {
        $Results |
            Export-Csv -Path $ReportPath `
                -NoTypeInformation -Encoding UTF8
    }

    if ($Errors.Count -gt 0) {
        $Errors |
            Set-Content -Path $ErrorLogPath -Encoding UTF8
    }

    Write-Host "`n========== SUMMARY ==========" -ForegroundColor Cyan
    Write-Host "Permanently deleted: $( @($Results | Where-Object Status -eq 'Permanently deleted').Count )"
    Write-Host "Failed:              $( @($Results | Where-Object Status -eq 'Failed').Count )"
    Write-Host "Report: $ReportPath"
    Write-Host "Error log: $ErrorLogPath"
}
