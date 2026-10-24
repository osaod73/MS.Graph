
# ============================================================
# Restore Deleted Microsoft 365 Groups by Email
# Input : C:\SchoolUsers\DeletedGroups.csv
# Column: Mail
# Output: RestoreGroupsReport.csv
#         GroupsNotMatched.csv
# ============================================================

$ErrorActionPreference = "Stop"

$CsvPath      = "C:\SchoolUsers\DeletedGroups.csv"
$ReportPath   = "C:\SchoolUsers\RestoreGroupsReport.csv"
$MissingPath  = "C:\SchoolUsers\GroupsNotMatched.csv"
$ErrorLogPath = "C:\SchoolUsers\RestoreGroupsErrors.txt"

$Results   = [System.Collections.Generic.List[object]]::new()
$Unmatched = [System.Collections.Generic.List[object]]::new()
$Errors    = [System.Collections.Generic.List[string]]::new()

function Get-GraphErrorDetails {
    param([System.Management.Automation.ErrorRecord]$Record)

    $details = $Record.Exception.Message

    if ($Record.ErrorDetails -and $Record.ErrorDetails.Message) {
        $details += " | Graph response: " + $Record.ErrorDetails.Message
    }

    return $details
}

try {
    # 1. Check input file
    if (-not (Test-Path -LiteralPath $CsvPath)) {
        throw "Input CSV file not found: $CsvPath"
    }

    $InputRows = @(Import-Csv -LiteralPath $CsvPath)

    if ($InputRows.Count -eq 0) {
        throw "The CSV file is empty."
    }

    if ($InputRows[0].PSObject.Properties.Name -notcontains "Mail") {
        throw "CSV must contain a column named Mail."
    }

    # Normalize emails
    $Emails = @(
        $InputRows |
        ForEach-Object {
            if ($null -ne $_.Mail) {
                $mail = ([string]$_.Mail).Trim()
                if ($mail.Length -gt 0) {
                    $mail.ToLowerInvariant()
                }
            }
        } |
        Sort-Object -Unique
    )

    if ($Emails.Count -eq 0) {
        throw "No valid email addresses were found in the CSV."
    }

    Write-Host "Email addresses to process: $($Emails.Count)" `
        -ForegroundColor Cyan

    # 2. Connect to Microsoft Graph
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    # Workaround for WAM window-handle problems
    $env:MSAL_USE_WAM = "false"

    Connect-MgGraph `
        -Scopes "Group.ReadWrite.All" `
        -UseDeviceCode `
        -NoWelcome `
        -ErrorAction Stop

    $context = Get-MgContext

    if (-not $context) {
        throw "Microsoft Graph connection could not be verified."
    }

    Write-Host "Connected to Microsoft Graph." -ForegroundColor Green

    # 3. Retrieve deleted groups with pagination
    # Use the typed deleted-items endpoint.
    $uri = "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.group?`$select=id,displayName,mail,mailNickname,groupTypes,securityEnabled,mailEnabled,proxyAddresses&`$top=999"

    $DeletedGroups = [System.Collections.Generic.List[object]]::new()

    while (-not [string]::IsNullOrWhiteSpace($uri)) {
        try {
            $page = Invoke-MgGraphRequest `
                -Method GET `
                -Uri $uri `
                -ErrorAction Stop
        }
        catch {
            $message = Get-GraphErrorDetails $_
            $Errors.Add("Retrieving deleted groups failed: $message")
            throw "Could not retrieve deleted groups. See $ErrorLogPath"
        }

        foreach ($item in $page.value) {
            $DeletedGroups.Add($item)
        }

        $uri = $page.'@odata.nextLink'
    }

    Write-Host "Deleted groups retrieved: $($DeletedGroups.Count)" `
        -ForegroundColor Cyan

    # 4. Match and restore
    foreach ($email in $Emails) {
        Write-Host "`nChecking: $email" -ForegroundColor Yellow

        # Match primary Mail first; then proxyAddresses as fallback.
        $matches = @(
            $DeletedGroups | Where-Object {
                $g = $_
                $primaryMatch = $false

                if (-not [string]::IsNullOrWhiteSpace([string]$g.mail)) {
                    $primaryMatch = $g.mail.Trim().Equals(
                        $email,
                        [System.StringComparison]::OrdinalIgnoreCase
                    )
                }

                $proxyMatch = $false
                if ($g.proxyAddresses) {
                    foreach ($address in $g.proxyAddresses) {
                        $normalizedAddress = ([string]$address).Trim()
                        $normalizedAddress = $normalizedAddress -replace '^(?i)smtp:', ''

                        if ($normalizedAddress.Equals(
                            $email,
                            [System.StringComparison]::OrdinalIgnoreCase
                        )) {
                            $proxyMatch = $true
                            break
                        }
                    }
                }

                $primaryMatch -or $proxyMatch
            }
        )

        if ($matches.Count -eq 0) {
            Write-Warning "No deleted group matched: $email"

            $Unmatched.Add([PSCustomObject]@{
                Mail   = $email
                Status = "Not found by Mail or proxyAddresses"
            })

            $Results.Add([PSCustomObject]@{
                Mail        = $email
                DisplayName = ""
                GroupId     = ""
                Status      = "Not found"
                Error       = "No matching deleted group found"
            })

            continue
        }

        if ($matches.Count -gt 1) {
            Write-Warning "Multiple deleted groups matched $email. Skipping."

            $Results.Add([PSCustomObject]@{
                Mail        = $email
                DisplayName = ($matches.displayName -join "; ")
                GroupId     = ($matches.id -join "; ")
                Status      = "Multiple matches - skipped"
                Error       = "Ambiguous email match; no group restored"
            })

            continue
        }

        $group = $matches[0]

        try {
            $restoreUri = "https://graph.microsoft.com/v1.0/directory/deletedItems/$($group.id)/restore"

            $restored = Invoke-MgGraphRequest `
                -Method POST `
                -Uri $restoreUri `
                -ErrorAction Stop

            $Results.Add([PSCustomObject]@{
                Mail        = $email
                DisplayName = $group.displayName
                GroupId     = $group.id
                Status      = "Restored"
                Error       = ""
            })

            Write-Host "RESTORED: $email" -ForegroundColor Green
        }
        catch {
            $message = Get-GraphErrorDetails $_
            $Errors.Add(
                "Restore failed | Mail=$email | GroupId=$($group.id) | $message"
            )

            $Results.Add([PSCustomObject]@{
                Mail        = $email
                DisplayName = $group.displayName
                GroupId     = $group.id
                Status      = "Failed"
                Error       = $message
            })

            Write-Host "RESTORE FAILED: $email" -ForegroundColor Red
            Write-Host $message -ForegroundColor Red
        }
    }
}
catch {
    $message = Get-GraphErrorDetails $_
    $Errors.Add("SCRIPT ERROR: $message")
    Write-Host "`nSCRIPT ERROR: $message" -ForegroundColor Red
}
finally {
    # 5. Export reports, including when an error occurs
    try {
        if ($Results.Count -gt 0) {
            $Results |
                Export-Csv -LiteralPath $ReportPath `
                    -NoTypeInformation -Encoding UTF8
        }

        if ($Unmatched.Count -gt 0) {
            $Unmatched |
                Export-Csv -LiteralPath $MissingPath `
                    -NoTypeInformation -Encoding UTF8
        }
        else {
            @() |
                Export-Csv -LiteralPath $MissingPath `
                    -NoTypeInformation -Encoding UTF8
        }

        if ($Errors.Count -gt 0) {
            $Errors |
                Set-Content -LiteralPath $ErrorLogPath -Encoding UTF8
        }

        Write-Host "`n========== SUMMARY ==========" -ForegroundColor Cyan
        Write-Host "Restored: $( @($Results | Where-Object Status -eq 'Restored').Count )"
        Write-Host "Failed:   $( @($Results | Where-Object Status -eq 'Failed').Count )"
        Write-Host "Not found: $( @($Results | Where-Object Status -eq 'Not found').Count )"
        Write-Host "Report: $ReportPath"
        Write-Host "Unmatched: $MissingPath"
        Write-Host "Error log: $ErrorLogPath"
    }
    catch {
        Write-Host "Could not export reports: $($_.Exception.Message)" `
            -ForegroundColor Red
    }
}
