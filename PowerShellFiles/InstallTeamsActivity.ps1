Install-Module Microsoft.Graph.Reports -Scope CurrentUser -Force -AllowClobber

Import-Module Microsoft.Graph.Reports

Get-Command -Module Microsoft.Graph.Reports *Teams*Activity*
