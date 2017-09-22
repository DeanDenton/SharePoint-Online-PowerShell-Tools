Function f_CSOM_SiteCol_GetProperyBags($spCtx) {
	Write-Host "----------------------------------------------------------------------------"  -foregroundcolor Green
	Write-Host "Reading PropertyBags values for $sSiteUrl !!" -ForegroundColor Green
	Write-Host "----------------------------------------------------------------------------"  -foregroundcolor Green

	$spoSiteCollection=$spCtx.Site
	$spCtx.Load($spoSiteCollection)
	$spoRootWeb=$spoSiteCollection.RootWeb
	$spCtx.Load($spoRootWeb)        
	$spoAllSiteProperties=$spoRootWeb.AllProperties
	$spCtx.Load($spoAllSiteProperties)
	$spCtx.ExecuteQuery()                
	$spoPropertyBagKeys=$spoAllSiteProperties.FieldValues.Keys
	#$spoPropertyBagKeys
	foreach($spoPropertyBagKey in $spoPropertyBagKeys){
		Write-Host "PropertyBag Key: " $spoPropertyBagKey " - PropertyBag Value: " $spoAllSiteProperties[$spoPropertyBagKey] -ForegroundColor Green
	}        
}
#################################
Function f_CSOM_SiteCol_GetAllFields ($spCtx) {
	$spFields = $spCtx.Web.Fields
	$spCtx.Load($spFields)
	$spCtx.ExecuteQuery()
	Return $spFields
}
#################################
Function f_CSOM_SiteCol_GetAllFields ($spCtx, $SiteColName) {
	$spFields = f_CSOM_SiteCol_GetAll -spCtx $spCtx
	$exists = $false
	foreach ($spItem in $spFields) {
		if ($spItem.InternalName -eq $SiteColName) {
			$exists = $true
		}
	}
	Return $exists
}
#################################
Function f_CSOM_SiteCol_Add ($spCtx, ) {
	If (!(f_CSOM_SiteCol_GetAllFields $spCtx)) {
		$spFields = f_CSOM_SiteCol_GetAll -spCtx $spCtx
		$fieldAsXML = "<Field Type='Text' 
				Name='"+$SiteColName+"'
				DisplayName='SPJ Name'
				Description='My Sample Site Column'
				ID='{545E5A0D-5553-4E02-97B3-F2D225B649D7}'
				Group='_Custom Site Column'
				Hidden='False'			  
				/>";
		
		$fieldOption = [Microsoft.SharePoint.Client.AddFieldOptions]::DefaultValue
		$field = $spFields.AddFieldAsXml($fieldAsXML, $true, $fieldOption)
		$spCtx.Load($field)
		try {
			$spCtx.ExecuteQuery()
			Write-Host "Site column" $SiteColName "created successfully"
		} catch {
			Write-Host "Error while creating site column" $SiteColName $_.Exception.Message 
		}
	}
}




<#
Function f_CSOM_SiteCol_GetUsage($spCtx) {
	$spCtx.Load($spUsage) 
    $spCtx.ExecuteQuery()  
    foreach($spSize in $spUsage){ 
        Write-Host ("Site Usage: " + $spSize.UsageInfo  )
    } 
}
#################################



$web = $scContext.Web
$siteCols = $web.Fields
$column = $sitecols.GetByInternalNameOrTitle("ColumnInternalName")


$scContext.load($web)
$scContext.load($siteCols)
$scContext.Load($column)
$scContext.ExecuteQuery()
#>