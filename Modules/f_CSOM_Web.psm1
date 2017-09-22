 ############

######################################
 # GET #
function f_CSOM_Web_Exists($webUrl){
        $spCtx = New-Object Microsoft.SharePoint.Client.ClientContext($webUrl)                     
        $spWeb = $spCtx.Web            
        $spCtx.Load($spWeb) 
        try {
            $spCtx.ExecuteQuery()
            return $true
        }
        catch {
            return $false
        }
    }
# # # # # # # # # # # # # # # # # # #
Function f_CSOM_Web_ShowSubSites($spCtx) {
    $rootWeb = $spCtx.Web
    $spSites = $rootWeb.Webs
    $spCtx.Load($rootWeb)
    $spCtx.Load($spSites)
    $spCtx.ExecuteQuery()

    foreach($spWeb in $spSites) {
        $spCtx.Load($spWeb)
        $spCtx.ExecuteQuery()
        Write-Host $spWeb.Title "-" $spWeb.Url
    }
}

# # # # # # # # # # # # # # # # # # #
Function f_CSOM_Web_ShowAll($webURL, $Login=$CtxLogin, $sPWD=$CtxPwd) {
    $spCtx =  f_CSOM_Ctx $webURL $Login $sPWD
    $rootWeb = $spCtx.Web  
    $spCtx.Load($rootWeb) 
    $spCtx.ExecuteQuery() 

    Write-Host  $rootWeb.Url "-" $rootWeb.Title 

    #Get SubSite
    $spSites = $spCtx.LoadQuery($rootWeb.Webs)
    $spCtx.ExecuteQuery()
    $spSites | % {
        f_CSOM_Web_ShowAll $_.Url  
     }
    $spCtx.dispose()
}
 # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Web_GetWelcomePage($spCtx) {
    $RootFolder = $spCtx.Web.RootFolder
    $spCtx.Load($RootFolder)
    try {
        $spCtx.ExecuteQuery()
        return $RootFolder.WelcomePage
    }
    catch {
        return "default.aspx"
    }
}

# # # # # # # # # # # # # # # # # # #
Function f_CSOM_Web_GetSvrRelUrl($spCtx){
    $spWeb = $spCtx.Web 
    $spCtx.Load($spWeb)
    try {
        $spCtx.ExecuteQuery()
	    Return $spWeb.ServerRelativeUrl
    }
    catch {
        return $NULL
    }
}

######################################
 # CREATE #
Function f_CSOM_Web_CreateSubSite($spCtx, $NewSubWebURLpart, $webTitle, $webTemplate) {
    $webInfo = New-Object Microsoft.SharePoint.Client.WebCreationInformation
    $webInfo.Url = $NewSubWebURLpart
    $webInfo.Title = $webTitle
    $webInfo.UseSamePermissionsAsParentSite = $true
    $webInfo.WebTemplate = $webTemplate
    $webInfo.Language = "1033"
    $spWeb = $spCtx.Web.Webs.Add($webInfo)

    $spCtx.Load($spWeb)
	Try {
	    $spCtx.ExecuteQuery()
		$Response = "Site Create Succeeded"
		}
	Catch {
		$Response = "Site Create FAILED"
		}
    Return $Response
}

######################################
 # DELETE #
Function f_CSOM_Web_Delete($webURL, $Login=$CtxLogin, $sPWD=$CtxPwd) {
    $spCtx = f_CSOM_Ctx $webURL $Login $sPWD
    $spWeb = $spCtx.Web
    $spCtx.Load($spWeb)
    $spWeb.DeleteObject()
    $spCtx.ExecuteQuery()
    $webTitle = $spWeb.Title
    Write-Host "Web Deleted: " $webTitle
}

# # # # # # # # # # # # # # # # # # #
Function f_CSOM_Web_Delete($spCtx) {
    $spWeb = $spCtx.Web
    $spCtx.Load($spWeb)
    $spWeb.DeleteObject()
	Try {
		$spCtx.ExecuteQuery()
		$Response = "Web Delete Succeeded"
	}
	Catch {
		$Response = "Web Delete Failed"
	}
	Return $Response
}



# # # # # # # # # # # # # # # # # # #
<#
Function f_CSOM_Web_Create($webUrl, $NewWebURL, $webTitle, $webTemplate, $Login=$CtxLogin, $sPWD=$CtxPwd) {
    $spCtx =  f_CSOM_Ctx $webURL $Login $sPWD
    $webInfo = New-Object Microsoft.SharePoint.Client.WebCreationInformation
    $webInfo.Url = $NewWebURL
    $webInfo.Title = $webTitle
    $webInfo.UseSamePermissionsAsParentSite = $true
    $webInfo.WebTemplate = $webTemplate
    $webInfo.Language = "1033"
    $spWeb = $spCtx.Web.Webs.Add($webInfo)

    $spCtx.Load($spWeb)
    $spCtx.ExecuteQuery()
    Write-Host "Web Created: " $webTitle
}
#>

<#
# # # # # # # # # # # # 
Function Get-AllProperties ($spCtx) {
    $spProperties = $spCtx.Web.AllProperties
    $spCtx.Load($spProperties)
    $spCtx.ExecuteQuery()
    ForEach( $spProperty in $spProperties){
        $spProperty | %{$_}
       # $spProperty | %{$_.FieldValues}
       # $spProperty | %{$_.Context}
        #$spProperty | %{$_.Path}
        #$spProperty | %{$_.TypedObject}
    }
}

# # # # # # # # # # 
Function Get-AllWebProperties ($spCtx) {
    $spWeb = $spCtx.Web
    $spCtx.Load($spWeb) 
    $spCtx.ExecuteQuery() 

    $spProperties = $spCtx.Web.AllProperties
    $spCtx.Load($spProperties)
    $spCtx.ExecuteQuery()

    ForEach( $spProperty in $spProperties){
        $spProperty | %{$_}
       # $spProperty | %{$_.FieldValues}
        $spProperty | %{$_.Context.Web}
        #$spProperty | %{$_.Path}
        #$spProperty | %{$_.TypedObject}
    }
}


#>