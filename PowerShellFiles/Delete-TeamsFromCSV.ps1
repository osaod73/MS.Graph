
# ==========================================
# Delete Microsoft Teams from CSV
# Confirm YES once to delete all teams
# ==========================================
# https://login.microsoft.com/device

$InputFile  = "C:\SchoolUsers\teams-to-delete.csv"
$ReportFile = "C:\SchoolUsers\teams-to-delete_Report.csv"

# Check CSV file
if (-not (Test-Path $InputFile)) {
    Write-Host "ERROR: Teams.csv not found: $InputFile" -ForegroundColor Red
    exit
}

# Import required modules
Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Teams
Import-Module Microsoft.Graph.Groups

# Read CSV
$Teams = @(Import-Csv -Path $InputFile)

if ($Teams.Count -eq 0 -or
    $Teams[0].PSObject.Properties.Name -notcontains "TeamId") {
    Write-Host "ERROR: CSV must contain a TeamId column and at least one row." -ForegroundColor Red
    exit
}

# Display teams to be deleted
Write-Host "`nTeams scheduled for deletion: $($Teams.Count)" -ForegroundColor Yellow
$Teams | ForEach-Object {
    Write-Host "Team ID: $($_.TeamId)"
}

# One-time confirmation
$Answer = Read-Host "`nType YES to delete ALL teams listed above"

if ($Answer -cne "YES") {
    Write-Host "Operation cancelled. No teams were deleted." -ForegroundColor Yellow
    exit
}

try {
    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "Team.ReadBasic.All", "Group.ReadWrite.All" -UseDeviceCode -ErrorAction Stop

    $Results = foreach ($Item in $Teams) {
        $TeamId = "$($Item.TeamId)".Trim()
        $TeamName = ""

        if ([string]::IsNullOrWhiteSpace($TeamId)) {
            [PSCustomObject]@{
                TeamId = $TeamId
                TeamName = ""
                Status = "Failed"
                Message = "Empty TeamId"
                Date = Get-Date
            }
            continue
        }

        try {
            # Verify team exists
            $Team = Get-MgTeam -TeamId $TeamId -ErrorAction Stop
            $TeamName = $Team.DisplayName

            Write-Host "`nProcessing: $TeamName ($TeamId)" -ForegroundColor Cyan

            # Delete associated Microsoft 365 group
            Remove-MgGroup -GroupId $TeamId -Confirm:$false -ErrorAction Stop

            Write-Host "Deletion requested successfully." -ForegroundColor Green

            [PSCustomObject]@{
                TeamId = $TeamId
                TeamName = $TeamName
                Status = "Deleted"
                Message = "Deletion request accepted"
                Date = Get-Date
            }
        }
        catch {
            Write-Host "FAILED: $TeamId - $($_.Exception.Message)" -ForegroundColor Red

            [PSCustomObject]@{
                TeamId = $TeamId
                TeamName = $TeamName
                Status = "Failed"
                Message = $_.Exception.Message
                Date = Get-Date
            }
        }
    }

    # Export results
    $Results | Export-Csv -Path $ReportFile -NoTypeInformation -Encoding UTF8

    Write-Host "`n===================================" -ForegroundColor Green
    Write-Host "Process completed." -ForegroundColor Green
    Write-Host "Report: $ReportFile" -ForegroundColor Yellow
}
catch {
    Write-Host "Connection or execution error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
