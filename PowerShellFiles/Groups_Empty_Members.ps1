
# ============================================================
# Report empty Microsoft 365 / Security Groups
# Microsoft Graph PowerShell
# ============================================================

$ErrorActionPreference = "Stop"

$Folder = "C:\SchoolUsers"
$Report = Join-Path $Folder "Empty_Groups_Report.csv"
$ErrorReport = Join-Path $Folder "Empty_Groups_Errors.csv"

if (-not (Test-Path $Folder)) {
    New-Item -Path $Folder -ItemType Directory -Force |
        Out-Null
}

# Install authentication module if missing
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication `
        -Scope CurrentUser -Force -AllowClobber
}

Import-Module Microsoft.Graph.Authentication

# Connect to Microsoft Graph
try {
    Connect-MgGraph -Scopes @(
        "Group.Read.All",
        "GroupMember.Read.All"
    ) -UseDeviceCode -NoWelcome -ErrorAction Stop
}
catch {
    Write-Host "Graph connection failed: $($_.Exception.Message)" `
        -ForegroundColor Red
    return
}

$EmptyGroups = [System.Collections.Generic.List[object]]::new()
$Errors = [System.Collections.Generic.List[object]]::new()

try {
    # Retrieve all groups with pagination
    $Groups = [System.Collections.Generic.List[object]]::new()

    $Uri = "https://graph.microsoft.com/v1.0/groups?`$select=id,displayName,mail,description,groupTypes,mailEnabled,securityEnabled,createdDateTime&`$top=999"

    while ($Uri) {
        $Page = Invoke-MgGraphRequest `
            -Method GET -Uri $Uri -OutputType PSObject

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
            -Activity "Checking empty groups" `
            -Status "$Index of $Total : $($Group.displayName)" `
            -PercentComplete (($Index / [math]::Max(1,$Total)) * 100)

        try {
            # Retrieve all direct members, including non-user objects
            $MemberUri = "https://graph.microsoft.com/v1.0/groups/$($Group.id)/members?`$select=id&`$top=999"

            $HasMembers = $false

            $MemberPage = Invoke-MgGraphRequest `
                -Method GET -Uri $MemberUri -OutputType PSObject

            if (@($MemberPage.value).Count -gt 0) {
                $HasMembers = $true
            }

            # Continue checking all pages to verify the membership
            $NextMemberUri = $MemberPage.'@odata.nextLink'

            while ($NextMemberUri) {
                $MemberPage = Invoke-MgGraphRequest `
                    -Method GET -Uri $NextMemberUri -OutputType PSObject

                if (@($MemberPage.value).Count -gt 0) {
                    $HasMembers = $true
                }

                $NextMemberUri = $MemberPage.'@odata.nextLink'
            }

            if (-not $HasMembers) {
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

                $EmptyGroups.Add([PSCustomObject]@{
                    GroupName       = $Group.displayName
                    GroupId         = $Group.id
                    Email           = $Group.mail
                    GroupType       = $GroupType
                    Description     = $Group.description
                    CreatedDateTime = $Group.createdDateTime
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

    Write-Progress -Activity "Checking empty groups" -Completed

    $EmptyGroups |
        Sort-Object GroupName |
        Export-Csv -Path $Report -NoTypeInformation -Encoding UTF8

    $Errors |
        Export-Csv -Path $ErrorReport -NoTypeInformation -Encoding UTF8

    Write-Host "`nReport completed successfully." -ForegroundColor Green
    Write-Host "Total groups checked: $Total"
    Write-Host "Empty groups found: $($EmptyGroups.Count)"
    Write-Host "Report: $Report"

    if ($Errors.Count -gt 0) {
        Write-Host "Groups with checking errors: $($Errors.Count)" `
            -ForegroundColor Yellow
        Write-Host "Error report: $ErrorReport"
    }

    $EmptyGroups |
        Format-Table GroupName, GroupType, Email, GroupId -AutoSize
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
