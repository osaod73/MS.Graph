
# ============================================================
# Report groups with more than 100 direct user members
# Microsoft Graph REST API
# Output: C:\SchoolUsers\Groups_Over_100_Members.csv
# ============================================================

$ErrorActionPreference = "Stop"

$Folder = "C:\SchoolUsers"
$Report = Join-Path $Folder "Groups_Over_100_Members.csv"
$ErrorReport = Join-Path $Folder "Groups_Over_100_Errors.csv"
$Threshold = 100

if (-not (Test-Path $Folder)) {
    New-Item -Path $Folder -ItemType Directory -Force | Out-Null
}

# Install authentication module if necessary
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication `
        -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Authentication

try {
    Connect-MgGraph -Scopes @(
        "Group.Read.All",
        "GroupMember.Read.All"
    ) -UseDeviceCode -NoWelcome -ErrorAction Stop
}
catch {
    Write-Host "Graph login failed: $($_.Exception.Message)" `
        -ForegroundColor Red
    return
}

$Results = [System.Collections.Generic.List[object]]::new()
$Errors = [System.Collections.Generic.List[object]]::new()

try {
    # Get all groups
    $Uri = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,description,mail,groupTypes,mailEnabled,securityEnabled&`$top=99"
    $Groups = [System.Collections.Generic.List[object]]::new()

    while ($Uri) {
        $Page = Invoke-MgGraphRequest -Method GET `
            -Uri $Uri -OutputType PSObject

        foreach ($Group in $Page.value) {
            $Groups.Add($Group)
        }

        $Uri = $Page.'@odata.nextLink'
    }

    $Total = $Groups.Count
    $Index = 0

    foreach ($Group in $Groups) {
        $Index++

        Write-Progress `
            -Activity "Checking group memberships" `
            -Status "$Index / $Total - $($Group.displayName)" `
            -PercentComplete (($Index / [math]::Max(1,$Total)) * 100)

        try {
            # Count only direct user members
            $MemberUri = "https://graph.microsoft.com/v1.0/groups/$($Group.id)/members/microsoft.graph.user/`$count"

            $Headers = @{
                ConsistencyLevel = "eventual"
            }

            $CountResponse = Invoke-MgGraphRequest `
                -Method GET `
                -Uri $MemberUri `
                -Headers $Headers `
                -OutputType PSObject

            $Count = [int64]$CountResponse

            if ($Count -gt $Threshold) {
                $Results.Add([PSCustomObject]@{
                    GroupName       = $Group.displayName
                    GroupId         = $Group.id
                    Email           = $Group.mail
                    Description     = $Group.description
                    DirectUserCount = $Count
                    GroupType       = if ($Group.groupTypes -contains "Unified") {
                        "Microsoft 365"
                    } elseif ($Group.securityEnabled) {
                        "Security"
                    } elseif ($Group.mailEnabled) {
                        "Distribution"
                    } else {
                        "Other"
                    }
                    MailEnabled     = $Group.mailEnabled
                    SecurityEnabled = $Group.securityEnabled
                })
            }
        }
        catch {
            $Errors.Add([PSCustomObject]@{
                GroupName = $Group.displayName
                GroupId   = $Group.id
                Error     = $_.Exception.Message
            })
        }
    }

    Write-Progress -Activity "Checking group memberships" -Completed

    $Results |
        Sort-Object DirectUserCount -Descending |
        Export-Csv -Path $Report -NoTypeInformation -Encoding UTF8

    $Errors |
        Export-Csv -Path $ErrorReport -NoTypeInformation -Encoding UTF8

    Write-Host "`nReport completed." -ForegroundColor Green
    Write-Host "Groups with more than $Threshold direct users: $($Results.Count)"
    Write-Host "Report: $Report"

    if ($Errors.Count -gt 0) {
        Write-Host "Groups that could not be checked: $($Errors.Count)" `
            -ForegroundColor Yellow
        Write-Host "Error report: $ErrorReport"
    }

    $Results |
        Format-Table GroupName, DirectUserCount, GroupType, Email -AutoSize
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
