$url = "https://maritzllc.sharepoint.com/sites/Dev-Dean03/"

$clientDll = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
$runtimeDll = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")

$cred = get-credential
$clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($url) 
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($cred.username, $cred.password) 
$clientContext.Credentials = $credentials 
if (!$clientContext.ServerObjectIsNull.Value) 
{ 
    Write-Host "Connected to SharePoint site: '$Url'" -ForegroundColor Green 
}
$clientContext.Site.Features.Add('9c0834e1-ba47-4d49-812b-7d4fb6fea211',$true,[Microsoft.SharePoint.Client.FeatureDefinitionScope]::None)
$clientContext.ExecuteQuery()