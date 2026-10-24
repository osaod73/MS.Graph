
# Restore deleted Microsoft 365 users using Mail
# Input : C:\SchoolUsers\DeletedUsers.csv
# Reports: C:\SchoolUsers\RestoreUsersReport.csv
#          C:\SchoolUsers\UsersNotMatched.csv

$ErrorActionPreference = "Stop"

$CsvPath       = "C:\SchoolUsers\DeletedUsers.csv"
$ReportPath    = "C:\SchoolUsers\RestoreUsersReport.csv"
$UnmatchedPath = "C:\SchoolUsers\UsersNotMatched.csv"

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

    # Avoid WAM interactive-window issues where supported
    $env:MSAL_USE_WAM = "false"

    Connect-MgGraph -Scopes "User.ReadWrite.All","User.DeleteRestore.All" -NoWelcome -UseDeviceCode

    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }

    $inputUsers = Import-Csv -Path $CsvPath

    if (-not ($inputUsers | Get-Member -Name Mail -MemberType NoteProperty)) {
        throw "CSV must contain a column named Mail."
    }

    # Retrieve deleted users, following all Graph result pages
    $deletedUsers = @(
        Get-MgDirectoryDeletedItemAsUser -All `
            -Property "id,displayName,mail,userPrincipalName,proxyAddresses,deletedDateTime"
    )

    $report = [System.Collections.Generic.List[object]]::new()
    $unmatched = [System.Collections.Generic.List[object]]::new()

    foreach ($entry in $inputUsers) {
        $email = ([string]$entry.Mail).Trim().ToLowerInvariant()

        if ([string]::IsNullOrWhiteSpace($email)) {
            $unmatched.Add([PSCustomObject]@{
                Mail   = $entry.Mail
                Reason = "Empty Mail value"
            })
            continue
        }

        # Match exact email values only
        $matches = @(
            $deletedUsers | Where-Object {
                $u = $_

                $values = @()
                if ($u.Mail) {
                    $values += ([string]$u.Mail).ToLowerInvariant()
                }
                if ($u.UserPrincipalName) {
                    $values += ([string]$u.UserPrincipalName).ToLowerInvariant()
                }
                foreach ($proxy in @($u.ProxyAddresses)) {
                    if ($proxy -match '^(?i)smtp:') {
                        $values += ($proxy -replace '^(?i)smtp:', '').ToLowerInvariant()
                    }
                }

                $values -contains $email
            }
        )

        if ($matches.Count -eq 0) {
            $unmatched.Add([PSCustomObject]@{
                Mail   = $entry.Mail
                Reason = "No exact email match among deleted users"
            })
            continue
        }

        if ($matches.Count -gt 1) {
            $unmatched.Add([PSCustomObject]@{
                Mail   = $entry.Mail
                Reason = "Multiple deleted accounts matched; skipped for safety"
            })
            continue
        }

        $user = $matches[0]

        try {
            # Restore the matching deleted user
            Restore-MgDirectoryDeletedItem `
                -DirectoryObjectId $user.Id `
                -ErrorAction Stop | Out-Null

            $report.Add([PSCustomObject]@{
                Mail              = $entry.Mail
                DisplayName       = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                ObjectId          = $user.Id
                DeletedDateTime   = $user.DeletedDateTime
                Status            = "Restored"
                Error             = ""
            })

            # Remove restored user from the in-memory deleted list
            $deletedUsers = @(
                $deletedUsers | Where-Object { $_.Id -ne $user.Id }
            )

            Write-Host "RESTORED: $email" -ForegroundColor Green
        }
        catch {
            $report.Add([PSCustomObject]@{
                Mail              = $entry.Mail
                DisplayName       = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                ObjectId          = $user.Id
                DeletedDateTime   = $user.DeletedDateTime
                Status            = "Failed"
                Error             = $_.Exception.Message
            })

            Write-Host "FAILED: $email - $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    $report | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    $unmatched | Export-Csv -Path $UnmatchedPath -NoTypeInformation -Encoding UTF8

    Write-Host "`nCompleted." -ForegroundColor Cyan
    Write-Host "Restoration report: $ReportPath"
    Write-Host "Unmatched report:   $UnmatchedPath"
    Write-Host "Deleted users scanned: $($deletedUsers.Count + $report.Count)"
}
catch {
    Write-Error "Script stopped: $($_.Exception.Message)"
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
