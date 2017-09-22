#f_CSOM_DOC.psm1
#Functions for Document Management

#        Add-Type -Path "G:\03 Docs\10 MVP\03 MVP Work\11 PS Scripts\Office 365\Microsoft.SharePoint.Client.DocumentManagement.dll"

 

#################################

Function f_CSOM_DOC_DocSetCreate ($spCtx, $DocLibName, $DocSetName, $CntTypeID="0x0120D520") {
    #https://gallery.technet.microsoft.com/office/How-to-create-a-document-3448341e
	$spDocLib=$spCtx.Web.Lists.GetByTitle($DocLibName)
	$spRootFolder=$spDocLib.RootFolder

	#Getting the Document Set Content Type by ID -> In this case we are using the default one in SPO
	$spDocSetCntType=$spCtx.Site.RootWeb.ContentTypes.GetById($CntTypeID)        
	$spCtx.Load($spDocSetCntType)                
	$spCtx.ExecuteQuery()

	#Creating the Document Set in the target Doc. Library
	$spDocSet=[Microsoft.SharePoint.Client.DocumentSet.DocumentSet]        
	$spDocSet::Create($spCtx,$spRootFolder,$DocSetName,$spDocSetCntType.Id)        
	$spCtx.ExecuteQuery()

    Return $spDocSet
}
