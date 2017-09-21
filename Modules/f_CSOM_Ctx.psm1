#Module/f_spCtx.ps1

Function f_CSOM_Ctx($CtxURL, $CtxLogin, $CtxPwd , $isSPO=$true) {
    $spCtx = New-Object Microsoft.SharePoint.Client.ClientContext($CtxURL)  -ErrorAction stop

    If ($isSPO) {
        $spCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($CtxLogin, $CtxPwd)   
        $spCtx.Credentials = $spCredentials  
    } Else {
        $spCtx.Credentials = New-Object System.Net.NetworkCredential($CtxLogin, $CtxPwd)
    }
    Return $spCtx    
} 
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Ctx_Get($CtxURL, $Credentials=$NULL, $getCredAlways=$false, $isSPO=$true) {
	If (($getCredAlways -eq $true) ){
		$Credentials = Get-Credential

	}
    $spCtx = New-Object Microsoft.SharePoint.Client.ClientContext($CtxURL)  -ErrorAction stop

	$CtxLogin = $Credentials.UserName
	$CtxPwd = $Credentials.Password 

    If ($isSPO) {
        $spCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($CtxLogin, $CtxPwd)   
        $spCtx.Credentials = $spCredentials  
    } Else {
        $spCtx.Credentials = New-Object System.Net.NetworkCredential($CtxLogin, $CtxPwd)
    }
    Return $spCtx    
}

Function f_CSOM_Ctx_TenantAdmin($TenantAdminURL) {
    $Credentials = Get-Credential
    $CtxLogin = $Credentials.UserName
	$CtxPwd = $Credentials.Password 

    $spCtx = New-Object Microsoft.SharePoint.Client.ClientContext($TenantAdminURL)  -ErrorAction stop
    $spCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($CtxLogin, $CtxPwd)   

    $spoTenant= New-Object Microsoft.Online.SharePoint.TenantAdministration.Tenant($spCtx)
    $spoTenantSiteCollections=$spoTenant.GetSiteProperties(0,$true)
    $spCtx.Load($spoTenantSiteCollections)
    $spCtx.ExecuteQuery()
    Return $spCtx    
} 