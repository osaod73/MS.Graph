
# ============================================================
# Report: All Microsoft 365 Users with Per-User MFA Status
# Output: CSV report
# ============================================================

$ErrorActionPreference = "Stop"

# Install Microsoft Graph module if not installed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users

try {
    # Connect to Microsoft Graph
  Connect-MgGraph -Scopes "User.Read.All", "Policy.Read.AuthenticationMethod" -UseDeviceCode -NoWelcome

    Write-Host "Retrieving users..." -ForegroundColor Cyan

    # Get all users
    $Users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,Department,JobTitle

    $Report = [System.Collections.Generic.List[object]]::new()
    $Counter = 0

    foreach ($User in $Users) {
        $Counter++

        Write-Progress `
            -Activity "Checking MFA status" `
            -Status "$Counter / $($Users.Count): $($User.UserPrincipalName)" `
            -PercentComplete (($Counter / $Users.Count) * 100)

        try {
            # Read per-user MFA state
            $Uri = "https://graph.microsoft.com/beta/users/$($User.Id)/authentication/requirements"

            $Result = Invoke-MgGraphRequest `
                -Method GET `
                -Uri $Uri `
                -ErrorAction Stop

            $MFAState = $Result.perUserMfaState

            if ([string]::IsNullOrWhiteSpace($MFAState)) {
                $MFAState = "Unknown"
            }

            $Report.Add([PSCustomObject]@{
                DisplayName       = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                Department        = $User.Department
                JobTitle          = $User.JobTitle
                MFAStatus         = $MFAState
            })
        }
        catch {
            $Report.Add([PSCustomObject]@{
                DisplayName       = $User.DisplayName
                UserPrincipalName = $User.UserPrincipalName
                Department        = $User.Department
                JobTitle          = $User.JobTitle
                MFAStatus         = "Error: $($_.Exception.Message)"
            })

            Write-Warning "Could not retrieve MFA for $($User.UserPrincipalName)"
        }
    }

    Write-Progress -Activity "Checking MFA status" -Completed

    # Export all users
    $OutputPath = Join-Path $env:USERPROFILE "Desktop\MFA_AllUsers_Report.csv"

    $Report |
        Sort-Object MFAStatus, DisplayName |
        Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    # Display summary
    Write-Host "`nMFA Report Summary" -ForegroundColor Green
    $Report | Group-Object MFAStatus |
        Select-Object @{Name="MFAStatus";Expression={$_.Name}},
                      @{Name="UserCount";Expression={$_.Count}} |
        Format-Table -AutoSize

    Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
}
catch {
    Write-Host "An error occurred: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
