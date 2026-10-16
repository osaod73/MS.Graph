
# Check Microsoft 365 User Accounts
# Input: Users.csv
# Output: User_Existence_Report.csv

# https://login.microsoft.com/device

$ErrorActionPreference = "Stop"

# تحديد مسار ملفات الإدخال والإخراج
if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $BasePath = $PSScriptRoot
}
else {
    $BasePath = (Get-Location).Path
}

$InputFile = Join-Path -Path $BasePath -ChildPath "Users.csv"

$OutputFile = Join-Path -Path $BasePath -ChildPath (
    "User_Existence_Report_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
)

Write-Host "Input file:  $InputFile"
Write-Host "Output file: $OutputFile"

$InputFile = "C:\SchoolUsers\Users.csv"
$OutputFile = "C:\SchoolUsers\User_Existence_Report.csv"


# Install module if required
if (-not (Get-Module -ListAvailable Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Users

try {
    if (-not (Test-Path $InputFile)) {
        throw "Users.csv was not found: $InputFile"
    }

    # Connect to Microsoft Graph
    Connect-MgGraph -Scopes "User.Read.All" -UseDeviceCode -NoWelcome

    $Users = Import-Csv $InputFile

    if (-not ($Users | Get-Member -Name UserPrincipalName -MemberType NoteProperty)) {
        throw "CSV must contain a UserPrincipalName column."
    }

    $Results = foreach ($Entry in $Users) {
        $UPN = ([string]$Entry.UserPrincipalName).Trim()

        if ([string]::IsNullOrWhiteSpace($UPN)) {
            continue
        }

        Write-Host "Checking: $UPN" -ForegroundColor Cyan

        try {
            # Search by exact user principal name
            $EscapedUPN = $UPN.Replace("'", "''")
            $FoundUser = Get-MgUser `
                -Filter "userPrincipalName eq '$EscapedUPN'" `
                -Property Id, DisplayName, UserPrincipalName, AccountEnabled `
                -ErrorAction Stop

            if ($FoundUser) {
                foreach ($User in $FoundUser) {
                    [PSCustomObject]@{
                        InputUPN         = $UPN
                        Exists           = "Yes"
                        DisplayName      = $User.DisplayName
                        UserPrincipalName = $User.UserPrincipalName
                        AccountEnabled   = $User.AccountEnabled
                        ObjectId         = $User.Id
                        Status           = if ($User.AccountEnabled) {
                            "Found - Enabled"
                        } else {
                            "Found - Disabled"
                        }
                        CheckedAt        = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                    }
                }
            }
            else {
                [PSCustomObject]@{
                    InputUPN          = $UPN
                    Exists            = "No"
                    DisplayName       = ""
                    UserPrincipalName = ""
                    AccountEnabled    = ""
                    ObjectId          = ""
                    Status            = "Not Found"
                    CheckedAt         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                }
            }
        }
        catch {
            [PSCustomObject]@{
                InputUPN          = $UPN
                Exists            = "Unknown"
                DisplayName       = ""
                UserPrincipalName = ""
                AccountEnabled    = ""
                ObjectId          = ""
                Status            = "Check Failed: $($_.Exception.Message)"
                CheckedAt         = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            }
        }
    }

    $Results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8

    Write-Host "`nReport created: $OutputFile" -ForegroundColor Green
    Write-Host "Total checked: $(@($Results).Count)"
    Write-Host "Existing: $( @($Results | Where-Object Exists -eq 'Yes').Count )"
    Write-Host "Not found: $( @($Results | Where-Object Exists -eq 'No').Count )"
    Write-Host "Check failed: $( @($Results | Where-Object Exists -eq 'Unknown').Count )"
}
catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
