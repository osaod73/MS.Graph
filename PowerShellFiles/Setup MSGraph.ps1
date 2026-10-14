#Install MsGraph
Install-Module Microsoft.Graph -Scope CurrentUser

#Connect to MsGraph
Connect-MgGraph -Scopes "User.Read.All","Group.ReadWrite.All"
