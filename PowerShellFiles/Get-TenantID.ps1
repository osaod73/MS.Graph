
# ==========================================
# Get Microsoft 365 Tenant ID
# Device Code Authentication
# ==========================================

$ErrorActionPreference = "Stop"

try {
    # Install required module if missing
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Install-Module Microsoft.Graph.Authentication `
            -Scope CurrentUser -Force -AllowClobber
    }

    Import-Module Microsoft.Graph.Authentication

    # Enter tenant domain
    $TenantDomain = Read-Host "Enter your tenant domain (example: contoso.com)"

    if ([string]::IsNullOrWhiteSpace($TenantDomain)) {
        throw "Tenant domain cannot be empty."
    }

    # Authenticate using device code
    Connect-MgGraph `
        -TenantId $TenantDomain `
        -Scopes "Organization.Read.All" `
        -UseDeviceCode `
        -NoWelcome

    # Get tenant information
    $Org = Invoke-MgGraphRequest `
        -Method GET `
        -Uri "https://graph.microsoft.com/v1.0/organization"

    $Result = foreach ($Item in $Org.value) {
        [PSCustomObject]@{
            OrganizationName = $Item.displayName
            TenantID         = $Item.id
            TenantDomain     = $TenantDomain
        }
    }

    $Result | Format-List

    New-Item -Path "C:\SchoolUsers" `
        -ItemType Directory -Force | Out-Null

    $Result | Export-Csv `
        "C:\SchoolUsers\TenantID_Report.csv" `
        -NoTypeInformation -Encoding UTF8

    Write-Host "Report saved to C:\SchoolUsers\TenantID_Report.csv" `
        -ForegroundColor Green
}
catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
}
