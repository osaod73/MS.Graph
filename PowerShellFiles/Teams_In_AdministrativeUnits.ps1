
# ============================================================
# Report: Microsoft Teams in Administrative Units
# Output: CSV
# ============================================================

$ErrorActionPreference = "Stop"

# Output file
$OutputFile = "C:\SchoolUsers\Teams_In_AdministrativeUnits.csv"

# Required Microsoft Graph scopes
$Scopes = @(
    "AdministrativeUnit.Read.All",
    "Group.Read.All",
    "Team.ReadBasic.All"
)

try {
    # Import required modules
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
    Import-Module Microsoft.Graph.Groups -ErrorAction Stop

    # Connect using device code to avoid WAM window issues
    Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome

    Write-Host "`nConnected to Microsoft Graph." -ForegroundColor Green

    # Get all administrative units
    $AdministrativeUnits = @(
        Get-MgDirectoryAdministrativeUnit -All `
            -Property "Id,DisplayName"
    )

    if ($AdministrativeUnits.Count -eq 0) {
        Write-Host "No administrative units were found." `
            -ForegroundColor Yellow
        return
    }

    $Results = [System.Collections.Generic.List[object]]::new()
    $Errors  = [System.Collections.Generic.List[object]]::new()

    # Cache teams already checked to avoid duplicate API calls
    $TeamCache = @{}

    foreach ($AU in $AdministrativeUnits) {

        Write-Host "`nAdministrative Unit: $($AU.DisplayName)" `
            -ForegroundColor Cyan

        try {
            # Get all direct members of this administrative unit
            $Members = @(
                Get-MgDirectoryAdministrativeUnitMember `
                    -AdministrativeUnitId $AU.Id -All
            )
        }
        catch {
            Write-Warning "Could not read members of $($AU.DisplayName)"

            $Errors.Add([PSCustomObject]@{
                AdministrativeUnitName = $AU.DisplayName
                AdministrativeUnitId   = $AU.Id
                Error                  = $_.Exception.Message
            })

            continue
        }

        foreach ($Member in $Members) {

            # Only Microsoft 365 groups can have an associated Team
            $Type = $Member.AdditionalProperties["@odata.type"]

            if ($Type -ne "#microsoft.graph.group") {
                continue
            }

            $GroupId = $Member.Id

            if ([string]::IsNullOrWhiteSpace($GroupId)) {
                continue
            }

            try {
                # Check each group only once
                if (-not $TeamCache.ContainsKey($GroupId)) {

                    $TeamCache[$GroupId] = $null

                    try {
                        $Team = Get-MgGroupTeam `
                            -GroupId $GroupId `
                            -ErrorAction Stop

                        if ($null -ne $Team) {
                            $Group = Get-MgGroup `
                                -GroupId $GroupId `
                                -Property "Id,DisplayName" `
                                -ErrorAction Stop

                            $TeamCache[$GroupId] = [PSCustomObject]@{
                                TeamName = $Group.DisplayName
                                TeamId   = $GroupId
                            }
                        }
                    }
                    catch {
                        # A group without an associated Team is skipped.
                        # Log errors other than common "not a Team" responses.
                        $Message = $_.Exception.Message

                        if (
                            $Message -notmatch '404|NotFound|Not Found|does not exist'
                        ) {
                            $Errors.Add([PSCustomObject]@{
                                AdministrativeUnitName = $AU.DisplayName
                                AdministrativeUnitId   = $AU.Id
                                GroupId                = $GroupId
                                Error                  = $Message
                            })
                        }
                    }
                }

                $TeamInfo = $TeamCache[$GroupId]

                if ($null -ne $TeamInfo) {

                    $Results.Add([PSCustomObject]@{
                        AdministrativeUnitName = $AU.DisplayName
                        AdministrativeUnitId   = $AU.Id
                        TeamName                = $TeamInfo.TeamName
                        TeamId                  = $TeamInfo.TeamId
                    })
                }
            }
            catch {
                $Errors.Add([PSCustomObject]@{
                    AdministrativeUnitName = $AU.DisplayName
                    AdministrativeUnitId   = $AU.Id
                    GroupId                = $GroupId
                    Error                  = $_.Exception.Message
                })
            }
        }
    }

    # Remove duplicate rows if the same Team is found more than once
    $FinalResults = @(
        $Results |
            Sort-Object AdministrativeUnitName, TeamName, TeamId -Unique
    )

    # Export report
    if ($FinalResults.Count -gt 0) {
        $FinalResults |
            Export-Csv -Path $OutputFile `
                -NoTypeInformation -Encoding UTF8

        Write-Host "`nReport created successfully:" `
            -ForegroundColor Green
        Write-Host $OutputFile -ForegroundColor Yellow
        Write-Host "Total rows: $($FinalResults.Count)"
    }
    else {
        Write-Warning "No Teams were found in the administrative units."
    }

    # Export errors separately when any occur
    if ($Errors.Count -gt 0) {

        $ErrorFile = Join-Path `
            ([Environment]::GetFolderPath("Desktop")) `
            "Teams_In_AdministrativeUnits_Errors.csv"

        $Errors |
            Export-Csv -Path $ErrorFile `
                -NoTypeInformation -Encoding UTF8

        Write-Warning "$($Errors.Count) errors occurred."
        Write-Host "Error report: $ErrorFile" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Script failed: $($_.Exception.Message)"
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
