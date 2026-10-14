
# 1. إغلاق جلسة Microsoft Graph الحالية
Disconnect-MgGraph -ErrorAction SilentlyContinue

# 2. تعطيل WAM لهذه الجلسة
$env:MSAL_USE_WAM = "false"

# 3. الاتصال باستخدام رمز الجهاز
Connect-MgGraph -Scopes "User.ReadWrite.All","Organization.Read.All" -UseDeviceCode

# 4. التحقق من نجاح الاتصال
Get-MgContext

# https://login.microsoft.com/device
