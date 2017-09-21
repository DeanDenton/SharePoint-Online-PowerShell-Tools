#$CSOM_Path   = "C:\SYNC\SharePoint\MSH Development - Code\dllCSOM\Microsoft.SharePointOnline.CSOM.16.1.6518.1200\lib\net45"
#Add-Type -Path "$CSOM_Path\Microsoft.SharePoint.Client.dll" 
    
$SiteNamePart = "Template01"
$NewSiteNamePart = $SiteNamePart+'a'
$TemplateSiteUrl = "https://maritzllc.sharepoint.com/sites/Dev-Dean03/"+$SiteNamePart
$NewSiteUrl = "https://maritzllc.sharepoint.com/sites/Dev-Dean03/"+$NewSiteNamePart

$outputPath = "C:\SYNC\DEV\Output\"

#If($Credentials -eq $Null) {
    $Credentials = Get-Credential
#}
Connect-PnPOnline -Url $TemplateSiteUrl -Credentials $Credentials
 
Get-PnPProvisioningTemplate -Out ($outputPath + $siteNamePart + "-PnPtemplate.xml") -PersistBrandingFiles
Get-PnPProvisioningTemplate -Out ($outputPath + $siteNamePart + "-PnPtemplate.pnp") -PersistBrandingFiles

#Create Site
Connect-PnPOnline -Url $NewSiteUrl -Credentials $Credentials

Apply-PnPProvisioningTemplate -Path ($outputPath + $siteNamePart + "-PnPtemplate.pnp") -ProvisionContentTypesToSubWebs 