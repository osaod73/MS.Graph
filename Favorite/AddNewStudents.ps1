# 1. إغلاق جلسة Microsoft Graph الحالية
Disconnect-MgGraph -ErrorAction SilentlyContinue

# 2. تعطيل WAM لهذه الجلسة
$env:MSAL_USE_WAM = "false"

# 3. الاتصال باستخدام رمز الجهاز
Connect-MgGraph -Scopes "User.ReadWrite.All","Organization.Read.All" -UseDeviceCode

# 4. التحقق من نجاح الاتصال
Get-MgContext

# https://login.microsoft.com/device

#Connect to MsGraph
#Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All"
#Connect-MgGraph -Scopes "Directory.AccessAsUser.All"

#List all licenses 
Get-MgSubscribedSku | Select-Object SkuPartNumber, SkuId

# Create new users from CSV
$i = 0
$NewUsers = Import-Csv -Path "C:\SchoolUsers\newusers.csv" -Encoding UTF8

foreach ($User in $NewUsers) {
    $UserPrincipalName = $User.UserPrincipalName
    $existingUser = Get-MgUser -UserId $UserPrincipalName -ErrorAction SilentlyContinue
    if ($existingUser) {
        Write-Host "User $UserPrincipalName already exists." -ForegroundColor Red
        continue
    }
    $PasswordProfile = @{
        password = $User.Password
        forceChangePasswordNextSignIn = $false
        forceChangePasswordNextSignInWithMfa = $false
    }

    # Build user object
    $userObject = @{
        accountEnabled    = $true
        displayName       = $User.DisplayName
        givenName         = $User.First
        surname           = $User.Surname
        jobTitle          = $User.JobTitle 
        department        = $User.Department
        country           = "Palestinian Authority"
        mailNickname      = $User.MailNickName
        userPrincipalName = $User.UserPrincipalName
        preferredLanguage = "ar-SA"
        usageLocation     = "PS"
        passwordProfile   = $PasswordProfile
        passwordPolicies  = "DisablePasswordExpiration"
    }
    $jsonBody = $userObject | ConvertTo-Json -Depth 5 -Compress
    try {
        $i++
        $user = Invoke-MgGraphRequest -Method POST -Uri "/v1.0/users" -Body $jsonBody -ContentType "application/json"
        $licenseParams = @{
            addLicenses = @(@{skuId = "314c4481-f395-4525-be8b-2ec4bb1e9d91"})
            removeLicenses = @()
        }
        Invoke-MgGraphRequest -Method POST -Uri "/v1.0/users/$($user.id)/assignLicense" -Body ($licenseParams | ConvertTo-Json)
        Write-Host ("User {0} has been created ... ({1}/{2})" -f $User.UserPrincipalName, $i ,$NewUsers.Count) -ForegroundColor Yellow
    }
    catch {
        Write-Host ("Failed to create the account for {0}." -f $User.DisplayName) -ForegroundColor Red

        if ($_.Exception.Response -and $_.Exception.Response.Content) {
            $errorContent = $_.Exception.Response.Content.ReadAsStringAsync().Result
            Write-Host "Graph API Error:" -ForegroundColor Red
            Write-Host $errorContent -ForegroundColor Red
        } else {
            Write-Host ("Error: {0}" -f $_.Exception.Message) -ForegroundColor Red
        }
    }
}

#Note: SkuId For students: 314c4481-f395-4525-be8b-2ec4bb1e9d91