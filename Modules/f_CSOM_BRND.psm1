#f_CSOM_BRND.psm1

<#
Function f_CSOM_BRND_ (){

}
#>
######################################
Function f_CSOM_BRND_MstrPgEditing ($ctx, $bAllow=$true){
	$ctx.Load($ctx.Site)
	$ctx.Load($ctx.Web.Webs)
	$ctx.Site.AllowMasterPageEditing=$bAllow

	$ctx.ExecuteQuery()
	$ctx.Load($ctx.Site)
	$ctx.ExecuteQuery()
}
######################################  
Function f_CSOM_BRND_ApplyTheme ($ctx, $palette, $font){
	$ctx.Load($ctx.Web)
	$ctx.Load($ctx.Web.ThemeInfo)
	$ctx.ExecuteQuery()
	$palette=$ctx.web.ServerRelativeUrl+$palette
	$font=$ctx.Web.ServerRelativeUrl+$font
	$mynu=Out-Null
	Write-Host "Current theme " $ctx.Web.ThemeInfo.AccessibleDescription
	Write-Host $ctx.Web.ThemeInfo.ThemeBackgroundImageUri
	$ctx.Web.ApplyTheme($palette,$font,$mynu, $true)
	$ctx.ExecuteQuery()
}
######################################
Function f_CSOM_BRND_SetTheme ($ctx, $ThemeURL, $FontSchemeURL, $BackgroundURL) {
    $web = $spCtx.Web
    $web.ApplyTheme( $ThemeURL , $FontSchemeURL , $BackgroundURL , $True)
    $web.update()
    $ctx.Load($web)
    $ctx.ExecuteQuery()
}
######################################    
Function f_CSOM_BRND_SetMasterPage ($spCtx, $SystemMasterPageUrl,$CustomMasterPageUrl) {
	$spWeb = $spCtx.Web

	#"{0}/_catalogs/masterpage/oslo.master", 
    $spWeb.MasterUrl = $SystemMasterPageUrl
    $spWeb.CustomMasterUrl = $CustomMasterPageUrl
    $spWeb.Update()
	#$spCtx.Load($spWeb)
    $spCtx.ExecuteQuery()

}
        
        
  
        

  
  
  
<#

  # Paths to SDK. Please verify location on your computer.
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 

# Insert the credentials and the name of the admin site
$Username="admin@tenant.onmicrosoft.com"
$AdminPassword=Read-Host -Prompt "Password" -AsSecureString
$AdminUrl="https://tenant.sharepoint.com/sites/powie1"



$paletteUrl="_catalogs/theme/15/Palette001.spcolor"
$fontUrl="_catalogs/theme/15/fontscheme001.spfont"


Set-SPoTheme -Username $Username -AdminPassword $AdminPassword -Url $AdminUrl -palette $paletteUrl -font $fontUrl
#>