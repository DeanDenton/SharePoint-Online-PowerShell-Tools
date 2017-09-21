$url = "https://maritzllc.sharepoint.com/teams/IT-LearnIT/"

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
$clientContext.Site.Features.Add('2a6bf8e8-10b5-42f2-9d3e-267dfb0de8d4',$true,[Microsoft.SharePoint.Client.FeatureDefinitionScope]::None)
$clientContext.ExecuteQuery()