Get-Module -Name Microsoft.Online.SharePoint.PowerShell -ListAvailable | Select Name,Version
Install-Module -Name Microsoft.Online.SharePoint.PowerShell
$adminUPN = Read-Host "what is your username?"
$orgName = Read-Host "what is your O365 tenant name"
$userCredential = Get-Credential -UserName $adminUPN -Message "Type the password."
Connect-SPOService -Url https://$orgName-admin.sharepoint.com -Credential $userCredential