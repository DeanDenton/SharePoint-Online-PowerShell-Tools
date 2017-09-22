#f_CSOM_CntType.psm1
<#
    Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Publishing.dll'
    Add-Type -Path 'C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll'

Function f_CSOM_CntType_() {
Get: 
Add: 
Edit: http://social.technet.microsoft.com/wiki/contents/articles/31444.sharepoint-online-content-types-in-powershell-edit.aspx

}
#>
#################################
Function f_CSOM_CntType_ListReorder($spCtx, $ListName, $ContentTypeNamesInOrder) {
	#$ContentTypesInOrder = "Content Type 1", "Content Type 2", "Content Type 3"
	$spList = $spCtx.Web.Lists.GetByTitle($ListName)
	$spCtx.Load($spList)
	$spCtx.ExecuteQuery()

	$spCntTypes = $spList.ContentTypes
	$spCtx.Load($spCntTypes)
	$spCtx.ExecuteQuery()

	$ctList = New-Object System.Collections.Generic.List[Microsoft.SharePoint.Client.ContentTypeId]
	Foreach($ct in $ContentTypeNamesInOrder){
		$ctToInclude = $spCntTypes | Where {$_.Name -eq $ct}
		$ctList.Add($ctToInclude.Id)
	}
	$spList.RootFolder.UniqueContentTypeOrder = $ctList
	$spList.Update()
	$spCtx.Load($spList)
	$spCtx.ExecuteQuery()
}
#################################
Function f_CSOM_CntType_GetAllByList ($spCtx, $ListName){
	$spListCntTypes = $spCtx.Web.Lists.GetByTitle($ListName).ContentTypes
	$spCtx.Load($spListCntTypes)
	$spCtx.ExecuteQuery()
	Return $spListCntTypes
	#foreach($spCntType in $spListCntTypes) { Write-Host $spCntType.Name }
}
#################################
Function f_CSOM_CntType_GetAllByList2 ($spCtx, $ListTitle){
	$spList=$spCtx.Web.Lists.GetByTitle($ListTitle)
	$spCtx.Load($spList)
	$spCtx.Load($spList.ContentTypes)
	$spCtx.ExecuteQuery()
	foreach($ct in $spList.ContentTypes) {
		Write-Host $ct.Name 
		Write-Host $ct.ID
		#Write-Host $ct.DisplayFormTemplateName
		if($ct.DisplayFormUrl -ne "") {
			Write-Host $ct.DisplayFormUrl
		}
	}
<#
$ctx.Web.ContentTypes
$ctx.Web.AvailableContentTypes
	Change Name
	$ct.DisplayFormTemplateName="DifferentForm"
	$ct.Update($false)
	$spCtx.ExecuteQuery()
	# # # # # # 
	Reset Default DisplayFormUrl
	foreach($cc in $ll.ContentTypes){
       if($ct.Name -eq "Item"){
         $ct.DisplayFormUrl=""
         $ct.Update($false)
        $spCtx.ExecuteQuery()
       }
    }
#>
}
#################################
Function f_CSOM_CntType_RemoveCTbyID ($ctx,$ContentTypeID) {
	$ctx.Load($ctx.Web)
	$ct=$ctx.Web.ContentTypes
	$ctx.Load($ct)
	$ctx.ExecuteQuery()
	$ctx.Web.ContentTypes.GetById($ContentTypeID).DeleteObject()
	$ctx.ExecuteQuery()
}



#################################
Function f_CSOM_CntType_UpdateListCT ($spCtx, $ListName, $ContentTypeName, $PropertyName, $PropertyValue) {
	$updateChildTypes = $false
	$spCntTypes = f_CSOM_CntType_GetAllByList -spCtx $spCtx -ListName $ListName
	foreach($spCntType in $spCntTypes) {   
		if($spCntType.Name -eq $ContentTypeName) {
			#Write-Host $cc.Name
			If ($PropertyName -eq "Description") {
				$spCntType.Description=$PropertyValue
				$spCntType.Update($updateChildType)
			}
			$ctx.ExecuteQuery()
		}     
	}
}
#################################
Function f_CSOM_CntType_UpdateListCT2 ($spCtx, $ListName, $CntTypeName, $DocTmplURL,$updateChildType=$true) {
#https://gallery.technet.microsoft.com/scriptcenter/Office-365-Update-Content-63bdad6b
    $web = $spCtx.Web
    $CntTypes = $web.ContentTypes
	$spCtx.Load($CntTypes);
	$spCtx.ExecuteQuery();

    foreach ($CntType in $CntTypes){  
        if($CntType.Name -eq $CntTypeName) {
            $CntType.DocumentTemplate = $DocTmplURL 
            $CntType.Update($updateChildType);
            $spCtx.Load($CntType);
	        $spCtx.ExecuteQuery();
           #Write-Host $ctype.Name "-" $ctype.DocumentTemplate    
        }
    }
}






#################################
Function f_CSOM_CntType_GetByGUID ($spCtx, $ContentTypeGUID) {
	$spCtx.Load($spCtx.Web)
	$spCntType=$spCtx.Web.ContentTypes.GetByID($ContentTypeGUID)
	$spCtx.Load($spCntType)
	$spCtx.ExecuteQuery()
	Return $spCntType
}

#################################
Function f_CSOM_CntType_UpdateCTbyGUID ($spCtx, $ContentTypeGUID,$PropertyName, $PropertyValue, $updateChildType=$true){
	$spCntType = f_CSOM_CntType_GetByGUID -spCtx $spCtx, -ContentTypeGUID $ContentTypeGUID
	If ($PropertyName -eq "Description") {
		$spCntType.Description=$PropertyValue
		$spCntType.Update($updateChildType)
	}
	$spCtx.ExecuteQuery()
}
#################################

Function f_CSOM_CntType_Create($spCtx,$CntTypeID,$CntTypeName,$SiteColName, $CntTypeGroupName ){
	#$CntTypeID = "0x0101"
	#$SiteColName = "My Custom Site Column"
	#$CntTypeName = "My Sample Content Type"
	#$CntTypeGroupName = "My Sample Content Types"
	#http://www.jussiroine.com/2014/11/recommended-approach-to-provisioning-content-types-in-sharepoint-2010-2013-and-sharepoint-online/

	# retrieve all site columns (fields)
	$spWeb = $spCtx.Web    
	$spFields = $spWeb.Fields         
	$spCtx.Load($spWeb)        
	$spCtx.Load($spFields)
	$spCtx.ExecuteQuery()

	# add a new content type - first, let's get all current ctypes
	$spWeb = $spCtx.Web      
	$spCntTypes = $spWeb.ContentTypes      
	$spCtx.Load($spWeb)     
	$spCtx.Load($spCntTypes)
	$spCtx.ExecuteQuery()

	# then load all content types that match the documents ID (0x0101)
	$spCntType = $spCntTypes.GetById($CntTypeID)
	$spCtx.Load($spCntType)
	$spCtx.ExecuteQuery()

	# create a new content type object
	$spNewCntTypeDef = New-Object Microsoft.SharePoint.Client.ContentTypeCreationInformation       
	$spNewCntTypeDef.Name = $CntTypeName  
	$spNewCntTypeDef.ParentContentType = $spCntType
	$spNewCntTypeDef.Group = $CntTypeGroupName  

	$spNewCntType = $contentTypes.Add($spNewCntTypeDef)
	$spCtx.ExecuteQuery()
	$spCtx.Load($spNewCntType)
	$spCtx.ExecuteQuery()

	Return $spNewCntType
}
#################################
Function f_CSOM_CntType_AddSiteCol ($spCtx, $CntTypeID, $siteColName, $updateChildType=$true) {
# finally add the custom site column to the content type
	$spWeb = $spCtx.Web
	$spField = $spWeb.Fields.GetByInternalNameOrTitle($siteColName) 
	$spCntType = $spWeb.ContentTypes.GetById($CntTypeID)
	$spCtx.Load($spWeb.Fields)
	$spCtx.Load($spField)
	$spCtx.Load($spWeb.ContentTypes)
	$spCtx.Load($spCntType.FieldLinks)
	$spCtx.Load($spCntType.Fields)
	
	$fieldRefLink = New-Object Microsoft.SharePoint.Client.FieldLinkCreationInformation
	$fieldRefLink.Field = $spField
	$spCntType.FieldLinks.Add($fieldRefLink)
	$spCntType.Update($updateChildType)
	$spCtx.ExecuteQuery()
}

#################################
  Function f_CSOM_CntType_ModFieldVsFildLinks($ctx, $ListTitle, $SiteColumn) {
	  $ctx.Load($ctx.Web.Lists)
	  $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
	  $ctx.Load($ll)
	  $ctx.Load($ll.ContentTypes)
	  $ctx.ExecuteQuery()
	  $field=$ctx.Web.Fields.GetByInternalNameOrTitle($SiteColumn)
	  foreach($cc in $ll.ContentTypes)
	  {
		 $link=new-object Microsoft.SharePoint.Client.FieldLinkCreationInformation
		 $link.Field=$field
		 $cc.FieldLinks.Add($link)
		 $cc.Update($false)
		 $ctx.ExecuteQuery()
	   }
   }




