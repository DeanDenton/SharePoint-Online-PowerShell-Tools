#f_CSOM_Tenant.psm1

Function f_CSOM_Tenant_Get ($adminUrl) {
#d2.Needs some thought on usage
	Connect-SPOService -Url $adminUrl
	$sites=(Get-SPOSite).Url
	foreach($site in $sites){
		Get-DeletedItems -Username $Username -AdminPassword $AdminPassword -Url $site
		$spCtx = f_CSOM_Ctx_Get -CtxURL $site
	}	
}