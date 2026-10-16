
# ============================================================
# Bulk removal of registered authentication methods
# Microsoft 365 / Microsoft Graph
# Input : Users.csv
# Output: MFA-Removal-Report.csv
# ============================================================
# https://login.microsoft.com/device

$ErrorActionPreference = "Stop"

# Paths relative to this script; fallback to current directory
$BasePath = if ($PSScriptRoot) {
    $PSScriptRoot
} else {
    (Get-Location).Path
}

$InputFile  = "C:\SchoolUsers\Users.csv"
$ReportFile = "C:\SchoolUsers\MFA-Removal-Report.csv"

# Install required module if missing
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser
}

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users

# Validate input
if (-not (Test-Path -LiteralPath $InputFile)) {
    throw "Input CSV not found: $InputFile"
}

$Users = @(Import-Csv -LiteralPath $InputFile)

if ($Users.Count -eq 0 -or
    $Users[0].PSObject.Properties.Name -notcontains "UserPrincipalName") {
    throw "CSV must contain a UserPrincipalName column and at least one user."
}

# Connect with delegated Graph permissions
Connect-MgGraph -Scopes @(
    "User.Read.All",
    "UserAuthenticationMethod.ReadWrite.All"
) -UseDeviceCode

# Authentication method type -> Graph endpoint collection
$EndpointMap = @{
    "#microsoft.graph.phoneAuthenticationMethod" =
        "phoneMethods"

    "#microsoft.graph.emailAuthenticationMethod" =
        "emailMethods"

    "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" =
        "microsoftAuthenticatorMethods"

    "#microsoft.graph.fido2AuthenticationMethod" =
        "fido2Methods"

    "#microsoft.graph.softwareOathAuthenticationMethod" =
        "softwareOathMethods"

    "#microsoft.graph.temporaryAccessPassAuthenticationMethod" =
        "temporaryAccessPassMethods"

    "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" =
        "windowsHelloForBusinessMethods"

    "#microsoft.graph.hardwareOathAuthenticationMethod" =
        "hardwareOathMethods"

    "#microsoft.graph.platformCredentialAuthenticationMethod" =
        "platformCredentialMethods"
}

$Queue   = [System.Collections.Generic.List[object]]::new()
$Results = [System.Collections.Generic.List[object]]::new()

# Phase 1: Read methods and build a preview
foreach ($Entry in $Users) {
    $UPN = ([string]$Entry.UserPrincipalName).Trim()

    if ([string]::IsNullOrWhiteSpace($UPN)) {
        $Results.Add([pscustomobject]@{
            UserPrincipalName = ""
            MethodType        = ""
            MethodId          = ""
            Status            = "Failed"
            Details           = "Empty UserPrincipalName"
            Time              = Get-Date
        })
        continue
    }

    try {
        $User = Get-MgUser -UserId $UPN -Property Id,UserPrincipalName

        $Uri = "https://graph.microsoft.com/v1.0/users/$($User.Id)/authentication/methods"
        $Response = Invoke-MgGraphRequest -Method GET -Uri $Uri

        foreach ($Method in $Response.value) {
            $Type = [string]$Method.'@odata.type'

            # Never delete the password authentication method
            if ($Type -eq "#microsoft.graph.passwordAuthenticationMethod") {
                continue
            }

            if ($EndpointMap.ContainsKey($Type)) {
                $Collection = $EndpointMap[$Type]

                $Queue.Add([pscustomobject]@{
                    UPN        = $UPN
                    UserId     = $User.Id
                    MethodType = $Type
                    MethodId   = [string]$Method.id
                    Collection = $Collection
                })
            }
            elseif ($Type -eq "#microsoft.graph.externalAuthenticationMethod") {
                $Queue.Add([pscustomobject]@{
                    UPN        = $UPN
                    UserId     = $User.Id
                    MethodType = $Type
                    MethodId   = [string]$Method.id
                    Collection = "externalAuthenticationMethods"
                })
            }
            else {
                $Results.Add([pscustomobject]@{
                    UserPrincipalName = $UPN
                    MethodType        = $Type
                    MethodId          = [string]$Method.id
                    Status            = "Skipped"
                    Details           = "Unsupported method type; not deleted"
                    Time              = Get-Date
                })
            }
        }
    }
    catch {
        $Results.Add([pscustomobject]@{
            UserPrincipalName = $UPN
            MethodType        = ""
            MethodId          = ""
            Status            = "Failed"
            Details           = $_.Exception.Message
            Time              = Get-Date
        })
    }
}

# Display preview
Write-Host "`nUsers in CSV: $($Users.Count)" -ForegroundColor Cyan
Write-Host "Methods queued for deletion: $($Queue.Count)" -ForegroundColor Yellow
Write-Host "Password methods are excluded." -ForegroundColor Green

$Queue |
    Group-Object UPN |
    Select-Object @{N="User";E={$_.Name}},
                  @{N="Methods";E={$_.Count}} |
    Format-Table -AutoSize

# Export preflight failures / unsupported types before confirmation
if ($Results.Count -gt 0) {
    $Results | Export-Csv -LiteralPath $ReportFile `
        -NoTypeInformation -Encoding UTF8
}

if ($Queue.Count -eq 0) {
    Write-Host "No supported methods were queued for deletion." -ForegroundColor Yellow
    return
}

# One confirmation for the entire batch
$Confirm = Read-Host "Type YES to delete ALL queued methods for these users"

if ($Confirm -cne "YES") {
    Write-Host "Cancelled. No queued methods were deleted." -ForegroundColor Yellow
    return
}

# Phase 2: Delete methods and log each result
foreach ($Item in $Queue) {
    try {
        $Uri = "https://graph.microsoft.com/v1.0/users/$($Item.UserId)/authentication/$($Item.Collection)/$($Item.MethodId)"

        # External authentication methods require /$ref
        if ($Item.Collection -eq "externalAuthenticationMethods") {
            $Uri += '/$ref'
        }

        Invoke-MgGraphRequest -Method DELETE -Uri $Uri | Out-Null

        $Status  = "Success"
        $Details = "Method deleted"
        Write-Host "[SUCCESS] $($Item.UPN) - $($Item.MethodType)" `
            -ForegroundColor Green
    }
    catch {
        $Status  = "Failed"
        $Details = $_.Exception.Message
        Write-Host "[FAILED] $($Item.UPN) - $Details" `
            -ForegroundColor Red
    }

    $Results.Add([pscustomobject]@{
        UserPrincipalName = $Item.UPN
        MethodType        = $Item.MethodType
        MethodId          = $Item.MethodId
        Status            = $Status
        Details           = $Details
        Time              = Get-Date
    })

    # Save progress after each deletion attempt
    $Results | Export-Csv -LiteralPath $ReportFile `
        -NoTypeInformation -Encoding UTF8
}

Write-Host "`nOperation completed." -ForegroundColor Cyan
Write-Host "Report: $ReportFile" -ForegroundColor Cyan

$Results |
    Group-Object Status |
    Select-Object Name, Count |
    Format-Table -AutoSize

Disconnect-MgGraph
