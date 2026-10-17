
# =====================================================
# Add Multiple Owners to Multiple Microsoft Teams
# Input : TeamsOwners.csv
# Output: TeamsOwners_Report.csv
# =====================================================

$ErrorActionPreference = "Stop"

# Files in the current working directory
$CsvPath = "C:\SchoolUsers\TeamsOwners.csv"
$ReportPath = "C:\SchoolUsers\TeamsOwners_Report.csv"

# Required Microsoft Graph modules
$RequiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Groups",
    "Microsoft.Graph.Teams"
)

try {
    # Install missing modules
    foreach ($ModuleName in $RequiredModules) {
        if (-not (Get-Module $ModuleName -ListAvailable)) {
            Install-Module $ModuleName -Scope CurrentUser `
                -Repository PSGallery -Force -AllowClobber
        }

        Import-Module $ModuleName -ErrorAction Stop
    }

    # Validate CSV file
    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }

    $Rows = @(Import-Csv -Path $CsvPath)

    if ($Rows.Count -eq 0) {
        throw "The CSV file contains no data."
    }

    if (-not ($Rows[0].PSObject.Properties.Name -contains "TeamId") -or
        -not ($Rows[0].PSObject.Properties.Name -contains "OwnerUPN")) {
        throw "CSV must contain TeamId and OwnerUPN columns."
    }

    # Validate rows
    foreach ($Row in $Rows) {
        if ([string]::IsNullOrWhiteSpace($Row.TeamId) -or
            [string]::IsNullOrWhiteSpace($Row.OwnerUPN)) {
            throw "A CSV row has an empty TeamId or OwnerUPN."
        }
    }

    # Sign in
    Connect-MgGraph -Scopes @(
        "Group.ReadWrite.All",
        "User.Read.All",
        "Team.ReadBasic.All"
    ) -UseDeviceCode -NoWelcome

    $Results = [System.Collections.Generic.List[object]]::new()

    # Cache teams and users to avoid repeated lookups
    $TeamCache = @{}
    $UserCache = @{}

    foreach ($Row in $Rows) {
        $TeamId = $Row.TeamId.Trim()
        $OwnerUPN = $Row.OwnerUPN.Trim()

        $Status = "Failed"
        $Details = ""

        try {
            # Get team and validate its existence
            if (-not $TeamCache.ContainsKey($TeamId)) {
                $TeamCache[$TeamId] = Get-MgTeam `
                    -TeamId $TeamId -ErrorAction Stop
            }

            # Get owner account
            if (-not $UserCache.ContainsKey($OwnerUPN)) {
                $UserCache[$OwnerUPN] = Get-MgUser `
                    -UserId $OwnerUPN -ErrorAction Stop
            }

            $Owner = $UserCache[$OwnerUPN]

            # Check existing owners
            $ExistingOwners = @(
                Get-MgGroupOwner -GroupId $TeamId -All -ErrorAction Stop
            )

            $AlreadyOwner = $ExistingOwners | Where-Object {
                $_.Id -eq $Owner.Id
            }

            if ($AlreadyOwner) {
                $Status = "Already Owner"
                $Details = "User is already an owner."
            }
            else {
                $Body = @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($Owner.Id)"
                }

                New-MgGroupOwnerByRef `
                    -GroupId $TeamId `
                    -BodyParameter $Body `
                    -ErrorAction Stop

                $Status = "Success"
                $Details = "Owner added successfully."
            }
        }
        catch {
            $Details = $_.Exception.Message
            Write-Host "FAILED: $OwnerUPN -> $TeamId : $Details" `
                -ForegroundColor Red
        }

        $Results.Add([PSCustomObject]@{
            TeamId   = $TeamId
            OwnerUPN = $OwnerUPN
            Status   = $Status
            Details  = $Details
            Date     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        })

        if ($Status -eq "Success") {
            Write-Host "SUCCESS: $OwnerUPN -> $TeamId" `
                -ForegroundColor Green
        }
        elseif ($Status -eq "Already Owner") {
            Write-Host "SKIPPED: $OwnerUPN is already an owner." `
                -ForegroundColor Yellow
        }
    }

    # Export report
    $Results | Export-Csv -Path $ReportPath `
        -NoTypeInformation -Encoding UTF8

    Write-Host "`nProcessing completed." -ForegroundColor Cyan
    Write-Host "Total rows: $($Results.Count)"
    Write-Host "Successful: $( @($Results | Where-Object Status -eq 'Success').Count )"
    Write-Host "Already owners: $( @($Results | Where-Object Status -eq 'Already Owner').Count )"
    Write-Host "Failed: $( @($Results | Where-Object Status -eq 'Failed').Count )"
    Write-Host "Report: $ReportPath" -ForegroundColor Cyan
}
catch {
    Write-Host "SCRIPT ERROR: $($_.Exception.Message)" `
        -ForegroundColor Red
}
finally {
    if (Get-MgContext -ErrorAction SilentlyContinue) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }
}
