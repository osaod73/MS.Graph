
# ============================================================
# Permanently delete all soft-deleted Entra ID groups
# WARNING: This action is irreversible.
# ============================================================

$ErrorActionPreference = "Stop"

$ReportPath = "C:\SchoolUsers\PermanentlyDeletedGroupsReport.csv"
$ErrorLogPath = "C:\SchoolUsers\PermanentlyDeleteGroupsErrors.txt"

$Results = [System.Collections.Generic.List[object]]::new()
$Errors = [System.Collections.Generic.List[string]]::new()

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # Use device-code authentication to avoid WAM window errors
    $env:MSAL_USE_WAM = "false"

    Connect-MgGraph `
        -Scopes "Group.ReadWrite.All" `
        -UseDeviceCode `
        -NoWelcome `
        -ErrorAction Stop

    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green

    # Retrieve deleted groups, following all pages
    $uri = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group?`$select=id,displayName,mail,mailNickname&`$top=999"

    $DeletedGroups = [System.Collections.Generic.List[object]]::new()

    while (-not [string]::IsNullOrWhiteSpace($uri)) {
        $response = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $uri `
            -ErrorAction Stop

        foreach ($group in $response.value) {
            $DeletedGroups.Add($group)
        }

        $uri = $response.'@odata.nextLink'
    }

    if ($DeletedGroups.Count -eq 0) {
        Write-Host "No deleted groups were found." -ForegroundColor Yellow
        return
    }

    Write-Host "`nDeleted groups found: $($DeletedGroups.Count)" `
        -ForegroundColor Yellow

    $DeletedGroups |
        Select-Object displayName, mail, mailNickname, id |
        Format-Table -AutoSize

    Write-Host "`nWARNING: Permanent deletion cannot be undone." `
        -ForegroundColor Red

    $confirmation = Read-Host `
        "Type YES to permanently delete ALL listed groups"

    if ($confirmation -cne "YES") {
        Write-Host "Cancelled. No groups were deleted." `
            -ForegroundColor Yellow
        return
    }

    foreach ($group in $DeletedGroups) {
        try {
            $deleteUri = "https://graph.microsoft.com/v1.0/directory/deletedItems/$($group.id)"

            Invoke-MgGraphRequest `
                -Method DELETE `
                -Uri $deleteUri `
                -ErrorAction Stop

            $Results.Add([PSCustomObject]@{
                DisplayName = $group.displayName
                Mail        = $group.mail
                GroupId     = $group.id
                Status      = "Permanently deleted"
                Error       = ""
            })

            Write-Host "DELETED: $($group.displayName) [$($group.id)]" `
                -ForegroundColor Green
        }
        catch {
            $message = $_.Exception.Message

            if ($_.ErrorDetails.Message) {
                $message += " | " + $_.ErrorDetails.Message
            }

            $Errors.Add(
                "GroupId=$($group.id); Mail=$($group.mail); Error=$message"
            )

            $Results.Add([PSCustomObject]@{
                DisplayName = $group.displayName
                Mail        = $group.mail
                GroupId     = $group.id
                Status      = "Failed"
                Error       = $message
            })

            Write-Host "FAILED: $($group.displayName)" `
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
    Write-Host "Successful deletions: $( @($Results | Where-Object Status -eq 'Permanently deleted').Count )"
    Write-Host "Failed operations:    $( @($Results | Where-Object Status -eq 'Failed').Count )"
    Write-Host "Report: $ReportPath"
    Write-Host "Error log: $ErrorLogPath"
}
