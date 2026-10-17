
# ============================================================
# Report Teams for ONE selected Administrative Unit
# Output: CSV on Desktop
# Authentication: Device Code
# ============================================================

$ErrorActionPreference = "Stop"

# Output files
$Desktop = [Environment]::GetFolderPath("Desktop")
$OutputFile = Join-Path $Desktop "Selected_AU_Teams_Report.csv"
$ErrorFile  = Join-Path $Desktop "Selected_AU_Teams_Errors.csv"

try {
    # 1. Import required modules
    $RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Identity.DirectoryManagement",
        "Microsoft.Graph.Groups",
        "Microsoft.Graph.Teams"
    )

    foreach ($Module in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $Module)) {
            Write-Host "Installing module: $Module" -ForegroundColor Yellow
            Install-Module $Module -Scope CurrentUser -Force -AllowClobber
        }

        Import-Module $Module -ErrorAction Stop
    }

    # 2. Connect to Microsoft Graph
    $Scopes = @(
        "AdministrativeUnit.Read.All",
        "Group.Read.All",
        "Team.ReadBasic.All"
    )

    Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome

    Write-Host "`nConnected successfully." -ForegroundColor Green

    # 3. Retrieve administrative units
    Write-Host "Loading administrative units..." -ForegroundColor Cyan

    $AUs = @(
        Get-MgDirectoryAdministrativeUnit -All |
        Sort-Object DisplayName
    )

    if ($AUs.Count -eq 0) {
        Write-Host "No administrative units found." -ForegroundColor Yellow
        return
    }

    # 4. Display selection menu
    Write-Host "`n========== ADMINISTRATIVE UNITS ==========" `
        -ForegroundColor Cyan

    for ($i = 0; $i -lt $AUs.Count; $i++) {
        Write-Host ("{0}. {1}" -f ($i + 1), $AUs[$i].DisplayName)
    }

    Write-Host "==========================================" `
        -ForegroundColor Cyan

    # 5. Ask the user to select one AU
    do {
        $Selection = Read-Host "`nEnter the number of the Administrative Unit"

        $Number = 0
        $ValidNumber = [int]::TryParse($Selection, [ref]$Number)

        if (
            -not $ValidNumber -or
            $Number -lt 1 -or
            $Number -gt $AUs.Count
        ) {
            Write-Host "Invalid selection. Please try again." `
                -ForegroundColor Red
        }
    } until (
        $ValidNumber -and
        $Number -ge 1 -and
        $Number -le $AUs.Count
    )

    $SelectedAU = $AUs[$Number - 1]

    Write-Host "`nSelected Unit: $($SelectedAU.DisplayName)" `
        -ForegroundColor Green
    Write-Host "AU ID: $($SelectedAU.Id)" -ForegroundColor Gray

    # 6. Retrieve direct members of the selected AU
    Write-Host "`nReading unit members..." -ForegroundColor Cyan

    $Members = @(
        Get-MgDirectoryAdministrativeUnitMember `
            -AdministrativeUnitId $SelectedAU.Id `
            -All
    )

    $Results = [System.Collections.Generic.List[object]]::new()
    $Errors  = [System.Collections.Generic.List[object]]::new()

    # 7. Check groups and identify Teams
    $GroupMembers = @(
        $Members | Where-Object {
            $_.AdditionalProperties["@odata.type"] -eq "#microsoft.graph.group"
        }
    )

    $TotalGroups = $GroupMembers.Count
    $Counter = 0

    foreach ($Member in $GroupMembers) {
        $Counter++
        $GroupId = $Member.Id

        Write-Progress `
            -Activity "Checking Microsoft Teams" `
            -Status "$Counter of $TotalGroups groups" `
            -PercentComplete $(if ($TotalGroups -gt 0) {
                ($Counter / $TotalGroups) * 100
            } else { 100 })

        try {
            # A group with an associated Team returns Team details
            $Team = Get-MgGroupTeam `
                -GroupId $GroupId `
                -ErrorAction Stop

            # Retrieve the group's display name
            $Group = Get-MgGroup `
                -GroupId $GroupId `
                -Property "Id,DisplayName" `
                -ErrorAction Stop

            $Results.Add([PSCustomObject]@{
                AdministrativeUnitName = $SelectedAU.DisplayName
                AdministrativeUnitId   = $SelectedAU.Id
                TeamName                = $Group.DisplayName
                TeamId                  = $Group.Id
            })

            Write-Host "Team found: $($Group.DisplayName)" `
                -ForegroundColor Green
        }
        catch {
            $Message = $_.Exception.Message

            # A 404 commonly means the group has no associated Team.
            # Record other errors for review.
            if ($Message -notmatch '404|NotFound|Not Found') {
                $Errors.Add([PSCustomObject]@{
                    GroupId = $GroupId
                    Error   = $Message
                })
            }
        }
    }

    Write-Progress -Activity "Checking Microsoft Teams" -Completed

    # 8. Export report
    $FinalResults = @(
        $Results | Sort-Object TeamName, TeamId -Unique
    )

    if ($FinalResults.Count -gt 0) {
        $FinalResults |
            Export-Csv -Path $OutputFile `
                -NoTypeInformation -Encoding UTF8

        Write-Host "`n========== REPORT COMPLETED ==========" `
            -ForegroundColor Green
        Write-Host "Administrative Unit: $($SelectedAU.DisplayName)"
        Write-Host "Total Teams: $($FinalResults.Count)"
        Write-Host "Report saved to: $OutputFile" `
            -ForegroundColor Yellow
    }
    else {
        Write-Host "`nNo Teams were found among the direct group members." `
            -ForegroundColor Yellow
    }

    # 9. Export errors separately
    if ($Errors.Count -gt 0) {
        $Errors |
            Export-Csv -Path $ErrorFile `
                -NoTypeInformation -Encoding UTF8

        Write-Host "`nSome groups could not be checked." `
            -ForegroundColor Yellow
        Write-Host "Review: $ErrorFile"
    }
}
catch {
    Write-Host "`nScript failed: $($_.Exception.Message)" `
        -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
