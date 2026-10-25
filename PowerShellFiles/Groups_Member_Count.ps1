
# ============================================================
# Report member counts for all Microsoft 365 groups
# Output: C:\SchoolUsers\Groups_Member_Count.csv
# ============================================================

$ErrorActionPreference = "Stop"

$Folder = "C:\SchoolUsers"
$Report = Join-Path $Folder "Groups_Member_Count.csv"
$ErrorReport = Join-Path $Folder "Groups_Member_Count_Errors.csv"

if (-not (Test-Path $Folder)) {
    New-Item -Path $Folder -ItemType Directory -Force | Out-Null
}

# Install Microsoft Graph Authentication module if required
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication `
        -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Authentication

# Connect to Graph using device code (avoids WAM window issues)
try {
    Connect-MgGraph -Scopes @(
        "Group.Read.All",
        "GroupMember.Read.All"
    ) -UseDeviceCode -NoWelcome -ErrorAction Stop
}
catch {
    Write-Host "Connection failed: $($_.Exception.Message)" `
        -ForegroundColor Red
    return
}

$Results = [System.Collections.Generic.List[object]]::new()
$Errors  = [System.Collections.Generic.List[object]]::new()

try {
    # Get all groups, following every page
    $Groups = [System.Collections.Generic.List[object]]::new()

    $Uri = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,mail,description,groupTypes,mailEnabled,securityEnabled,createdDateTime&`$top=999"

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
            -Activity "Counting group members" `
            -Status "$Index of $Total : $($Group.displayName)" `
            -PercentComplete (($Index / [math]::Max(1,$Total)) * 100)

        try {
            # Get direct members with pagination
            $Count = 0
            $MemberUri = "https://graph.microsoft.com/v1.0/groups/$($Group.id)/members?`$select=id&`$top=999"

            while ($MemberUri) {
                $MemberPage = Invoke-MgGraphRequest `
                    -Method GET -Uri $MemberUri -OutputType PSObject

                $Count += @($MemberPage.value).Count
                $MemberUri = $MemberPage.'@odata.nextLink'
            }

            $GroupType = if ($Group.groupTypes -contains "Unified") {
                "Microsoft 365"
            }
            elseif ($Group.securityEnabled -and $Group.mailEnabled) {
                "Mail-enabled Security"
            }
            elseif ($Group.securityEnabled) {
                "Security"
            }
            elseif ($Group.mailEnabled) {
                "Distribution"
            }
            else {
                "Other"
            }

            $Results.Add([PSCustomObject]@{
                GroupName       = $Group.displayName
                GroupId         = $Group.id
                Email           = $Group.mail
                GroupType       = $GroupType
                MemberCount     = $Count
                Description     = $Group.description
                CreatedDateTime = $Group.createdDateTime
                MailEnabled     = $Group.mailEnabled
                SecurityEnabled = $Group.securityEnabled
            })
        }
        catch {
            $Errors.Add([PSCustomObject]@{
                GroupName = $Group.displayName
                GroupId   = $Group.id
                Error     = $_.Exception.Message
            })
        }
    }

    Write-Progress -Activity "Counting group members" -Completed

    $Results |
        Sort-Object MemberCount -Descending |
        Export-Csv -Path $Report -NoTypeInformation -Encoding UTF8

    if ($Errors.Count -gt 0) {
        $Errors |
            Export-Csv -Path $ErrorReport -NoTypeInformation -Encoding UTF8
    }

    Write-Host "`nReport completed successfully." -ForegroundColor Green
    Write-Host "Total groups: $Total"
    Write-Host "Successfully counted: $($Results.Count)"
    Write-Host "Groups with errors: $($Errors.Count)"
    Write-Host "Report: $Report"

    $Results |
        Format-Table GroupName, GroupType, MemberCount, Email -AutoSize
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
