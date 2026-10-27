
# Restore all deleted Microsoft 365 / Entra ID users
# Report: C:\SchoolUsers\RestoreAllDeletedUsersReport.csv

$ErrorActionPreference = "Stop"

$ReportFolder = "C:\SchoolUsers"
$ReportPath = Join-Path $ReportFolder "RestoreAllDeletedUsersReport.csv"

# Create report folder if it does not exist
if (-not (Test-Path $ReportFolder)) {
    New-Item -Path $ReportFolder -ItemType Directory -Force | Out-Null
}

# Install required module if missing
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
}

# Use REST requests through Graph SDK authentication
Import-Module Microsoft.Graph.Authentication

# Avoid WAM interactive browser problems in embedded terminals
$env:MSAL_USE_WAM = "false"

$Scopes = @(
    "User.Read.All",
    "User.DeleteRestore.All"
)

try {
    Connect-MgGraph -Scopes $Scopes -UseDeviceCode -NoWelcome
}
catch {
    Write-Host "Sign-in failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Try signing in from a normal PowerShell terminal." -ForegroundColor Yellow
    return
}

$Results = [System.Collections.Generic.List[object]]::new()

try {
    # Get all deleted user objects, handling pagination
    $DeletedUsers = @(
        Invoke-MgGraphRequest `
            -Method GET `
            -Uri "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,deletedDateTime&`$top=100" `
            -OutputType PSObject |
        ForEach-Object { $_.value }
    )

    # Follow additional pages
    $NextLink = $null

    # Retrieve the initial response again to preserve @odata.nextLink
    $Page = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,deletedDateTime&`$top=100" `
        -OutputType PSObject

    $DeletedUsers = @($Page.value)

    while ($Page.'@odata.nextLink') {
        $Page = Invoke-MgGraphRequest `
            -Method GET `
            -Uri $Page.'@odata.nextLink' `
            -OutputType PSObject

        $DeletedUsers += @($Page.value)
    }

    Write-Host "Deleted users found: $($DeletedUsers.Count)" -ForegroundColor Cyan

    if ($DeletedUsers.Count -eq 0) {
        Write-Host "No deleted users found." -ForegroundColor Yellow
    }
    else {
        $Confirmation = Read-Host "Restore ALL $($DeletedUsers.Count) deleted users? Type YES"

        if ($Confirmation -cne "YES") {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return
        }

        foreach ($User in $DeletedUsers) {
            $Status = "FAILED"
            $ErrorMessage = ""

            try {
                # Restore the exact object by its unique ID
                $RestoreUri = "https://graph.microsoft.com/v1.0/directory/deletedItems/$($User.id)/restore"

                Invoke-MgGraphRequest `
                    -Method POST `
                    -Uri $RestoreUri `
                    -Body @{} `
                    -ContentType "application/json" `
                    -ErrorAction Stop | Out-Null

                $Status = "RESTORED"
                Write-Host "RESTORED: $($User.userPrincipalName) - $($User.displayName)" -ForegroundColor Green
            }
            catch {
                $ErrorMessage = $_.Exception.Message
                Write-Host "FAILED: $($User.userPrincipalName) - $ErrorMessage" -ForegroundColor Red
            }

            $Results.Add([PSCustomObject]@{
                DisplayName       = $User.displayName
                UserPrincipalName = $User.userPrincipalName
                Mail              = $User.mail
                DeletedDateTime   = $User.deletedDateTime
                ObjectId          = $User.id
                Status            = $Status
                Error              = $ErrorMessage
            })
        }
    }
}
catch {
    Write-Host "ERROR retrieving deleted users: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if ($Results.Count -gt 0) {
        $Results | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report saved to: $ReportPath" -ForegroundColor Cyan

        $Results | Group-Object Status |
            Select-Object Name, Count |
            Format-Table -AutoSize
    }

    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
