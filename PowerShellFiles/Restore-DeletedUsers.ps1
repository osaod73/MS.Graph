
# ==========================================================
# Restore Deleted Microsoft 365 Users from CSV
# Authentication: Device Code (avoids WAM window issue)
# CSV column: UserPrincipalName or Id
# ==========================================================

$CsvPath    = "C:\SchoolUsers\Users.csv"
$ReportPath = "C:\SchoolUsers\RestoreUsersReport.csv"

$ErrorActionPreference = "Stop"

try {
    # Create working folder if it does not exist
    $Folder = Split-Path $CsvPath -Parent

    if (-not (Test-Path $Folder)) {
        New-Item -Path $Folder -ItemType Directory -Force |
            Out-Null
    }

    # Install required Graph modules if missing
    $RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Identity.DirectoryManagement"
    )

    foreach ($Module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            Write-Host "Installing module: $Module" -ForegroundColor Cyan

            Install-Module $Module `
                -Scope CurrentUser -Force -AllowClobber
        }

        Import-Module $Module -ErrorAction Stop
    }

    # Check CSV file
    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }

    $InputUsers = @(Import-Csv -Path $CsvPath)

    if ($InputUsers.Count -eq 0) {
        throw "CSV file is empty."
    }

    $Columns = @($InputUsers[0].PSObject.Properties.Name)

    if (
        ($Columns -notcontains "UserPrincipalName") -and
        ($Columns -notcontains "Id")
    ) {
        throw "CSV must contain UserPrincipalName or Id."
    }

    # Connect to Microsoft Graph using device code
    Write-Host "`nConnecting to Microsoft Graph..." `
        -ForegroundColor Cyan

    Connect-MgGraph `
        -Scopes "User.ReadWrite.All" `
        -UseDeviceCode `
        -NoWelcome

    $Context = Get-MgContext

    if (-not $Context) {
        throw "Microsoft Graph connection was not established."
    }

    Write-Host "Connected successfully." -ForegroundColor Green

    # Retrieve deleted users
    Write-Host "`nReading deleted users..." -ForegroundColor Cyan

    $DeletedUsers = @(
        Get-MgDirectoryDeletedItemAsUser -All `
            -Property "id,displayName,userPrincipalName"
    )

    Write-Host "Deleted users found: $($DeletedUsers.Count)" `
        -ForegroundColor Cyan

    $Results = [System.Collections.Generic.List[object]]::new()
    $Queue = [System.Collections.Generic.List[object]]::new()

    # Match CSV users against deleted users
    foreach ($Row in $InputUsers) {
        $UPN = ""
        $ObjectId = ""

        if ($Columns -contains "UserPrincipalName") {
            $UPN = ([string]$Row.UserPrincipalName).Trim()
        }

        if ($Columns -contains "Id") {
            $ObjectId = ([string]$Row.Id).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($UPN) -and
            [string]::IsNullOrWhiteSpace($ObjectId)) {

            $Results.Add([PSCustomObject]@{
                UserPrincipalName = $UPN
                DisplayName       = ""
                ObjectId          = ""
                Status            = "Skipped"
                Message           = "UPN and Id are empty"
            })
            continue
        }

        $Match = $null

        # Object ID takes priority if supplied
        if (-not [string]::IsNullOrWhiteSpace($ObjectId)) {
            $Match = $DeletedUsers |
                Where-Object { $_.Id -eq $ObjectId } |
                Select-Object -First 1
        }
        elseif (-not [string]::IsNullOrWhiteSpace($UPN)) {
            $Match = $DeletedUsers |
                Where-Object {
                    $_.UserPrincipalName -eq $UPN
                } |
                Select-Object -First 1
        }

        if ($Match) {
            $AlreadyQueued = @(
                $Queue | Where-Object { $_.Id -eq $Match.Id }
            ).Count -gt 0

            if (-not $AlreadyQueued) {
                $Queue.Add([PSCustomObject]@{
                    Id                = $Match.Id
                    UserPrincipalName = $Match.UserPrincipalName
                    DisplayName       = $Match.DisplayName
                })
            }
        }
        else {
            $Results.Add([PSCustomObject]@{
                UserPrincipalName = $UPN
                DisplayName       = ""
                ObjectId          = $ObjectId
                Status            = "NotFound"
                Message           = "User not found in deleted users"
            })
        }
    }

    # Display users before restoration
    Write-Host "`nUsers scheduled for restoration:" `
        -ForegroundColor Yellow

    if ($Queue.Count -gt 0) {
        $Queue |
            Format-Table DisplayName, UserPrincipalName, Id -AutoSize

        $Confirmation = Read-Host `
            "Type YES to restore ALL listed users"

        if ($Confirmation -ceq "YES") {
            foreach ($User in $Queue) {
                try {
                    Restore-MgDirectoryDeletedItem `
                        -DirectoryObjectId $User.Id `
                        -ErrorAction Stop | Out-Null

                    $Results.Add([PSCustomObject]@{
                        UserPrincipalName = $User.UserPrincipalName
                        DisplayName       = $User.DisplayName
                        ObjectId          = $User.Id
                        Status            = "Restored"
                        Message           = "Restoration request succeeded"
                    })

                    Write-Host "RESTORED: $($User.UserPrincipalName)" `
                        -ForegroundColor Green
                }
                catch {
                    $Results.Add([PSCustomObject]@{
                        UserPrincipalName = $User.UserPrincipalName
                        DisplayName       = $User.DisplayName
                        ObjectId          = $User.Id
                        Status            = "Failed"
                        Message           = $_.Exception.Message
                    })

                    Write-Host "FAILED: $($User.UserPrincipalName)" `
                        -ForegroundColor Red

                    Write-Host $_.Exception.Message `
                        -ForegroundColor Red
                }
            }
        }
        else {
            foreach ($User in $Queue) {
                $Results.Add([PSCustomObject]@{
                    UserPrincipalName = $User.UserPrincipalName
                    DisplayName       = $User.DisplayName
                    ObjectId          = $User.Id
                    Status            = "Cancelled"
                    Message           = "Confirmation not provided"
                })
            }

            Write-Host "Restoration cancelled." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "No matching deleted users found." `
            -ForegroundColor Yellow
    }

    # Export results
    $Results |
        Export-Csv -Path $ReportPath `
            -NoTypeInformation -Encoding UTF8

    Write-Host "`n========================================"
    Write-Host "Operation completed."
    Write-Host "Report: $ReportPath" -ForegroundColor Cyan
    Write-Host "========================================"
}
catch {
    Write-Host "`nERROR: $($_.Exception.Message)" `
        -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
