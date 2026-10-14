
# ============================================================
# Microsoft 365 - Bulk User Update
# Source: CSV
# Existing users only
# Empty fields are ignored
# Logs: CSV + TXT
# ============================================================
# https://login.microsoft.com/device

$ErrorActionPreference = "Stop"

# Configuration
$BasePath = "C:\SchoolUsers"
$CsvPath  = Join-Path $BasePath "UsersUpdate.csv"
$LogPath  = Join-Path $BasePath "UpdateUsers_Log.txt"
$ReportPath = Join-Path $BasePath "UpdateUsers_Report.csv"

# Create folder if missing
if (-not (Test-Path $BasePath)) {
    New-Item -Path $BasePath -ItemType Directory -Force |
        Out-Null
}

# Logging function
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )

    $Line = "{0} [{1}] {2}" -f `
        (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), `
        $Level, $Message

    Add-Content -Path $LogPath -Value $Line `
        -Encoding UTF8

    switch ($Level) {
        "ERROR" {
            Write-Host $Line -ForegroundColor Red
        }
        "SUCCESS" {
            Write-Host $Line -ForegroundColor Green
        }
        "WARNING" {
            Write-Host $Line -ForegroundColor Yellow
        }
        default {
            Write-Host $Line -ForegroundColor Cyan
        }
    }
}

# Start a fresh log
"Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" |
    Set-Content -Path $LogPath -Encoding UTF8

try {
    # Validate input file
    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }

    # Install Graph module if missing
    if (-not (Get-Module -ListAvailable `
        -Name Microsoft.Graph.Users)) {

        Write-Log "Installing Microsoft.Graph.Users module"

        Install-Module Microsoft.Graph.Users `
            -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Users

    # Connect to Graph
    Write-Log "Connecting to Microsoft Graph"

    Connect-MgGraph `
        -Scopes "User.ReadWrite.All" `
        -UseDeviceCode `
        -ErrorAction Stop

    # Import CSV
    $Users = @(Import-Csv -Path $CsvPath -Encoding UTF8)

    if ($Users.Count -eq 0) {
        throw "CSV file contains no user records"
    }

    # Validate required column
    if (-not ($Users[0].PSObject.Properties.Name `
        -contains "UserPrincipalName")) {

        throw "Missing required CSV column: UserPrincipalName"
    }

    $Results = [System.Collections.Generic.List[object]]::new()

    $Total = $Users.Count
    $Updated = 0
    $Skipped = 0
    $Failed = 0

    $Index = 0

    foreach ($User in $Users) {
        $Index++

        $UPN = [string]$User.UserPrincipalName

        Write-Progress `
            -Activity "Updating Microsoft 365 users" `
            -Status "$Index of $Total : $UPN" `
            -PercentComplete (($Index / $Total) * 100)

        $ChangedFields = [System.Collections.Generic.List[string]]::new()
        $Status = "Skipped"
        $Details = ""

        try {
            # Validate account identifier
            if ([string]::IsNullOrWhiteSpace($UPN)) {
                throw "UserPrincipalName is empty"
            }

            # Find existing user only
            $ExistingUser = Get-MgUser `
                -UserId $UPN `
                -Property "id,userPrincipalName,displayName" `
                -ErrorAction Stop

            if ($null -eq $ExistingUser) {
                $Skipped++
                $Details = "User not found"
                Write-Log "$UPN - user not found; skipped" "WARNING"

                $Results.Add([PSCustomObject]@{
                    Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    UserPrincipalName = $UPN
                    Status = $Status
                    ChangedFields = ""
                    Details = $Details
                })
                continue
            }

            # Build updates only from non-empty CSV fields
            $Params = @{
                UserId = $ExistingUser.Id
                ErrorAction = "Stop"
            }

            $FieldMap = @{
                GivenName   = "GivenName"
                Surname     = "Surname"
                DisplayName = "DisplayName"
                Department  = "Department"
                JobTitle    = "JobTitle"
            }

            foreach ($CsvField in $FieldMap.Keys) {
                $Value = [string]$User.$CsvField

                # Empty or whitespace fields are never sent to Graph
                if (-not [string]::IsNullOrWhiteSpace($Value)) {
                    $Value = $Value.Trim()
                    $GraphField = $FieldMap[$CsvField]
                    $Params[$GraphField] = $Value
                    $ChangedFields.Add($CsvField)
                }
            }

            # No non-empty values to update
            if ($ChangedFields.Count -eq 0) {
                $Status = "Skipped"
                $Skipped++
                $Details = "No non-empty fields to update"

                Write-Log "$UPN - no fields to update" "WARNING"
            }
            else {
                # Update existing user
                Update-MgUser @Params

                $Status = "Success"
                $Updated++
                $Details = "Updated successfully"

                Write-Log (
                    "{0} - updated: {1}" -f
                    $UPN, ($ChangedFields -join ", ")
                ) "SUCCESS"
            }
        }
        catch {
            $Status = "Failed"
            $Failed++
            $Details = $_.Exception.Message

            Write-Log "$UPN - $Details" "ERROR"
        }

        $Results.Add([PSCustomObject]@{
            Date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            UserPrincipalName = $UPN
            Status = $Status
            ChangedFields = $ChangedFields -join ", "
            Details = $Details
        })
    }

    Write-Progress -Activity "Updating Microsoft 365 users" `
        -Completed

    # Export detailed report
    $Results | Export-Csv `
        -Path $ReportPath `
        -NoTypeInformation `
        -Encoding UTF8

    Write-Log "========================================"
    Write-Log "Processing completed"
    Write-Log "Total records: $Total"
    Write-Log "Updated successfully: $Updated"
    Write-Log "Skipped: $Skipped"
    Write-Log "Failed: $Failed"
    Write-Log "Text log: $LogPath"
    Write-Log "CSV report: $ReportPath"
}
catch {
    Write-Log "FATAL ERROR: $($_.Exception.Message)" "ERROR"
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
