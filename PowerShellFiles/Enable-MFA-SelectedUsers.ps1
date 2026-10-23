# ============================================================
# Microsoft 365 - Enable MFA for Selected Users
# CSV -> Security Group -> Conditional Access -> Enable MFA
# Does NOT install or update modules while running.
# ============================================================

$ErrorActionPreference = "Stop"

# ---------------- CONFIGURATION ----------------

$CsvPath = "C:\SchoolUsers\MFAUsers.csv"
$OutputFolder = "C:\SchoolUsers"
$ReportPath = Join-Path $OutputFolder "MFA-Results.csv"

$GroupName = "CA-Require-MFA-Selected-Users"
$PolicyName = "CA-Require-MFA-Selected-Users"

# ---------------- PREPARATION ----------------

try {
    if (-not (Test-Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force |
            Out-Null
    }

    if (-not (Test-Path $CsvPath)) {
        throw "CSV file not found: $CsvPath"
    }

    $Users = @(
        Import-Csv -Path $CsvPath |
        Where-Object {
            -not [string]::IsNullOrWhiteSpace(
                $_.UserPrincipalName
            )
        }
    )

    if ($Users.Count -eq 0) {
        throw "No valid users found in CSV."
    }

    $RequiredModules = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Users",
        "Microsoft.Graph.Groups",
        "Microsoft.Graph.Identity.SignIns"
    )

    foreach ($ModuleName in $RequiredModules) {
        if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
            throw "Module missing: $ModuleName. Install it in a separate PowerShell session."
        }

        Import-Module $ModuleName -ErrorAction Stop
    }

    Write-Host "CSV users found: $($Users.Count)" `
        -ForegroundColor Cyan
}
catch {
    Write-Host "PREPARATION ERROR: $($_.Exception.Message)" `
        -ForegroundColor Red
    return
}

# ---------------- CONNECT TO GRAPH ----------------

try {
    Connect-MgGraph -Scopes @(
        "User.Read.All",
        "Group.ReadWrite.All",
        "Policy.Read.All",
        "Policy.ReadWrite.ConditionalAccess"
    ) -UseDeviceCode -NoWelcome

    if (-not (Get-MgContext)) {
        throw "Microsoft Graph connection failed."
    }
}
catch {
    Write-Host "CONNECTION ERROR: $($_.Exception.Message)" `
        -ForegroundColor Red
    return
}

try {
    # ---------------- FIND OR CREATE SECURITY GROUP ----------------

    $Groups = @(
        Get-MgGroup -Filter "displayName eq '$GroupName'" -All
    )

    if ($Groups.Count -gt 1) {
        throw "Multiple groups named '$GroupName' found."
    }

    if ($Groups.Count -eq 1) {
        $Group = $Groups[0]

        if (-not $Group.SecurityEnabled -or
            $Group.GroupTypes -contains "DynamicMembership") {
            throw "Existing group is not a suitable static security group."
        }

        Write-Host "Using existing group: $GroupName" `
            -ForegroundColor Yellow
    }
    else {
        $Group = New-MgGroup `
            -DisplayName $GroupName `
            -Description "Selected users targeted by MFA Conditional Access" `
            -MailEnabled:$false `
            -MailNickname "CARequireMFASelectedUsers" `
            -SecurityEnabled:$true

        Write-Host "Security group created." `
            -ForegroundColor Green
    }

    # ---------------- LOAD EXISTING MEMBERS ----------------

    $ExistingMemberIds = @(
        Get-MgGroupMember -GroupId $Group.Id -All |
        ForEach-Object { $_.Id }
    )

    # ---------------- ADD USERS AND BUILD REPORT ----------------

    $Results = @(
        foreach ($Entry in $Users) {
            $UPN = $Entry.UserPrincipalName.Trim()

            try {
                $DirectoryUser = Get-MgUser `
                    -UserId $UPN `
                    -Property Id,DisplayName,UserPrincipalName

                if ($ExistingMemberIds -contains $DirectoryUser.Id) {
                    $Status = "Already a member"
                }
                else {
                    New-MgGroupMemberByRef `
                        -GroupId $Group.Id `
                        -BodyParameter @{
                            "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($DirectoryUser.Id)"
                        }

                    $ExistingMemberIds += $DirectoryUser.Id
                    $Status = "Added to group"
                }

                Write-Host "$UPN : $Status" -ForegroundColor Green

                [PSCustomObject]@{
                    UserPrincipalName = $UPN
                    DisplayName       = $DirectoryUser.DisplayName
                    Status            = $Status
                    Details           = ""
                }
            }
            catch {
                Write-Host "$UPN : FAILED - $($_.Exception.Message)" `
                    -ForegroundColor Red

                [PSCustomObject]@{
                    UserPrincipalName = $UPN
                    DisplayName       = ""
                    Status            = "Failed"
                    Details           = $_.Exception.Message
                }
            }
        }
    )

    $Results | Export-Csv -Path $ReportPath `
        -NoTypeInformation -Encoding UTF8

    $FailedUsers = @(
        $Results | Where-Object { $_.Status -eq "Failed" }
    )

    if ($FailedUsers.Count -gt 0) {
        throw "$($FailedUsers.Count) user(s) failed. Policy creation or activation stopped. Review $ReportPath."
    }

    # ---------------- FIND EXISTING POLICY ----------------

    $Policies = @(
        Get-MgIdentityConditionalAccessPolicy -All |
        Where-Object { $_.DisplayName -eq $PolicyName }
    )

    if ($Policies.Count -gt 1) {
        throw "Multiple policies named '$PolicyName' found. No policy changes made."
    }

    if ($Policies.Count -eq 1) {
        $Policy = $Policies[0]

        # Never modify a policy that does not target this group
        if (@($Policy.Conditions.Users.IncludeGroups) -notcontains $Group.Id) {
            throw "Existing policy does not target the expected group. No policy changes made."
        }

        if (@($Policy.GrantControls.BuiltInControls) -notcontains "mfa") {
            throw "Existing policy does not require MFA. No policy changes made."
        }

        if (@($Policy.Conditions.Applications.IncludeApplications) -notcontains "All") {
            throw "Existing policy does not target all applications. Review it manually; no changes made."
        }

        Write-Host ""
        Write-Host "Existing MFA policy found." -ForegroundColor Cyan
        Write-Host "Current state: $($Policy.State)"
        Write-Host "Policy ID: $($Policy.Id)"

        if ($Policy.State -eq "enabled") {
            Write-Host "MFA policy is already enabled." `
                -ForegroundColor Green
        }
        else {
            $Confirmation = Read-Host "Type YES to enable MFA enforcement"

            if ($Confirmation -cne "YES") {
                Write-Host "Activation cancelled."
                return
            }

            Update-MgIdentityConditionalAccessPolicy `
                -ConditionalAccessPolicyId $Policy.Id `
                -BodyParameter @{ state = "enabled" }

            Write-Host "Existing MFA policy enabled." `
                -ForegroundColor Green
        }
    }
    else {
        # ---------------- CREATE AND ENABLE NEW POLICY ----------------

        Write-Host ""
        Write-Host "A new Conditional Access policy will be created."
        Write-Host "Target group: $GroupName"
        Write-Host "Requirement: MFA for all cloud applications."

        $Confirmation = Read-Host "Type YES to create and ENABLE MFA"

        if ($Confirmation -cne "YES") {
            Write-Host "Cancelled. No policy created."
            return
        }

        $PolicyBody = @{
            displayName = $PolicyName
            state = "enabled"

            conditions = @{
                users = @{
                    includeUsers = @()
                    includeGroups = @($Group.Id)
                    excludeUsers = @()
                    excludeGroups = @()
                    excludeRoles = @()
                }

                applications = @{
                    includeApplications = @("All")
                    excludeApplications = @()
                }

                clientAppTypes = @("all")
            }

            grantControls = @{
                operator = "OR"
                builtInControls = @("mfa")
            }
        }

        $Policy = New-MgIdentityConditionalAccessPolicy `
            -BodyParameter $PolicyBody

        Write-Host "New MFA policy created and enabled." `
            -ForegroundColor Green
    }

    # ---------------- VERIFY FINAL STATE ----------------

    $FinalPolicy = Get-MgIdentityConditionalAccessPolicy `
        -ConditionalAccessPolicyId $Policy.Id

    Write-Host ""
    Write-Host "========================================"
    Write-Host "MFA OPERATION COMPLETED"
    Write-Host "========================================"
    Write-Host "Group:       $GroupName"
    Write-Host "Group ID:    $($Group.Id)"
    Write-Host "Policy:      $($FinalPolicy.DisplayName)"
    Write-Host "Policy ID:   $($FinalPolicy.Id)"
    Write-Host "State:       $($FinalPolicy.State)"
    Write-Host "CSV report:  $ReportPath"

    if ($FinalPolicy.State -ne "enabled") {
        Write-Host "WARNING: Policy is not enabled." `
            -ForegroundColor Yellow
    }
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" `
        -ForegroundColor Red
    Write-Host "Check the report: $ReportPath"
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}