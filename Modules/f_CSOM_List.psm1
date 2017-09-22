#$queryXML = '<View><Query><OrderBy><FieldRef Name="Modified" Ascending="FALSE"/></OrderBy></Query><RowLimit>1</RowLimit></View>'
#$queryXML = '<View><Query><Where><Eq><FieldRef Name="SiteLookup_x003A_Site_x0020_Type"/><Value Type="Text">Web</Value></Eq></Where></Query></View>'
#$siteHiddenLists = "appdata,Cache Profiles,Content and Structure Reports,Content type publishing error log,Converted Forms,Device Channels,Form Templates,List Template Gallery,Long Running Operation Status,Maintenance Log Library,Notification List,Quick Deploy Items,Relationships List,Solution Gallery,Style Library,Suggested Content Browser Locations,TaxonomyHiddenList,Theme Gallery,Translation Packages,Translation Status,User Information List,Variation Labels,Web Part Gallery,wfpub"
#$webHiddenLists = "fpdatasources,Composed Looks,wfsvc"
#$filterLists = $siteHiddenLists + "," + $webHiddenLists

# CREATE
Function f_CSOM_List_CreateNewByName($spCtx, $ListTitle, $ListTemplateName){
	#TemplateType https://waelmohamed.wordpress.com/2013/06/11/table-of-list-template-ids-and-splisttemplatetype-enumenration-members-in-sharepoint/
	$spWeb = $spCtx.Web
	$spCtx.Load($spWeb)
	$spCtx.ExecuteQuery()
	$listinfo =New-Object Microsoft.SharePoint.Client.ListCreationInformation
	$listinfo.Title = $ListTitle
	$listinfo.TemplateType = [Microsoft.SharePoint.Client.ListTemplateType] $ListTemplateName
	$spList = $spWeb.Lists.Add($listinfo)
	$spCtx.ExecuteQuery()
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_CreateNewByID($spCtx, $ListTitle, $ListTemplateID){
	$spWeb = $spCtx.Web
	$spCtx.Load($spWeb)
	$spCtx.ExecuteQuery()
	$listinfo =New-Object Microsoft.SharePoint.Client.ListCreationInformation
	$listinfo.Title = $ListTitle
	$listinfo.TemplateType =  $ListTemplateID
	$spList = $spWeb.Lists.Add($listinfo)
	$spCtx.ExecuteQuery()
}

# # # # # # # # # # # # # # # # # # # #
#GET
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_ExistsByTitle($spCtx, $listTitle) { 
	$spWeb = $spCtx.Web
	$spCtx.Load($spWeb)
    $spCtx.Load($spWeb.Lists)
	$spCtx.ExecuteQuery()
	$spList = $spWeb.Lists | Where-object {$_.title -eq $listTitle}
    If ($spList) {
        Return $true
    }
    Else {
		Return $false
    }
}


# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_GetMaxField ($spCtx,$listName,$fieldName) {
    $queryXML = '<View><Query><OrderBy><FieldRef Name="'+$fieldName+'" Ascending="FALSE"/></OrderBy></Query><RowLimit>1</RowLimit></View>'
    $spItem = f_CSOM_Item_GetAllByQuery -spCtx $spCtx -listTitle $listName -queryXML $queryXML
    Return  $spItem[$fieldName] 
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_GetAllInWeb($spCtx) { 
	$spLists = $spCtx.Web.Lists
	$spCtx.Load($spLists)
	$spCtx.ExecuteQuery()
    Return $spLists
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_Count($spCtx, $visibleOnly=$true) { 
	$spLists = $spCtx.Web.Lists
	$spCtx.Load($spLists)
	$spCtx.ExecuteQuery()
    $count = 0
    ForEach ($spList in $spLists) {
        If ($visibleOnly -eq $true) {
            If ($spList.Hidden -eq $false) {
                $count ++
            } 
        }
        Else {
            $count ++
        }
    }
    Return $count
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Lists_FullItemCount($spCtx, $visibleOnly=$true) { 
	$spLists = $spCtx.Web.Lists
	$spCtx.Load($spLists)
	$spCtx.ExecuteQuery()
    $count = 0
    ForEach ($spList in $spLists) {
        If ($visibleOnly -eq $true) {
            If ($spList.Hidden -eq $false) {
                $count += $spList.ItemCount
            } 
        }
        Else {
            $count += $spList.ItemCount
        }
    }
    Return $count
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_MaxItemCount($spCtx, $visibleOnly=$true) { 
	$spLists = $spCtx.Web.Lists
	$spCtx.Load($spLists)
	$spCtx.ExecuteQuery()
    $MaxCount = 0
    ForEach ($spList in $spLists) {
        If ($MaxCount -lt $spList.ItemCount) { 
            If ($visibleOnly -eq $true) {
                If ($spList.Hidden -eq $false) {
                    $MaxCount = $spList.ItemCount
                }
            }
            Else {
                $MaxCount = $spList.ItemCount
            }
        }
    }
    Return $MaxCount
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_GetByTitle($spCtx, $listTitle) { 
	$spList = $spCtx.Web.Lists.GetByTitle($listTitle)
	$spCtx.Load($spList)
	$spCtx.ExecuteQuery()
    Return $spList
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_LastModified ($spCtx, $spList) {
    If ($spList.ItemCount -eq 0){
        Return ""
    }
    Else {
        $query = New-Object Microsoft.SharePoint.Client.CamlQuery
        $query.ViewXml = '<View><Query><OrderBy><FieldRef Name="Modified" Ascending="FALSE"/></OrderBy></Query><RowLimit>1</RowLimit></View>'
        $spItems = $spList.GetItems($query)
        $spCtx.Load($spItems)
        $spCtx.ExecuteQuery()
        $LastModified = $spItems[0]["Modified"]
        $OutputString = $OutputString +$delimiter+$LastModified
        Return $LastModified
    }
}
# # # # # # # # # # # # # # # # # # # #

#EDIT ##########

# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_Rename($spCtx, $listTitle, $listTitleNew) { 
	$spList = $spCtx.Web.Lists.GetByTitle($listTitle)
	$spList.Title = $listTitleNew
	$spList.Update()
	$spCtx.ExecuteQuery()
    Return "List Renamed to $listTitleNew"
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_GetList_ContentTypes ($spCtx,$spList) {
    $spContentTypes = $spList.ContentTypes
    $spCtx.Load($spContentTypes)  
	$spCtx.ExecuteQuery() 
    Return $spContentTypes
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_SetProp_Attachments($spCtx, $listTitle, $bEnable) { 
	$spList=$spCtx.Web.Lists.GetByTitle($listTitle)
	$spList.EnableAttachments = $bAttachments 
	$spList.Update()
	$spCtx.ExecuteQuery()
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_FirstCreated ($spCtx, $spList) {
    If ($spList.ItemCount -eq 0){
        Return ""
    }
    Else {
        $query = New-Object Microsoft.SharePoint.Client.CamlQuery
        $query.ViewXml = '<View><Query><OrderBy><FieldRef Name="Created" Ascending="TRUE"/></OrderBy></Query><RowLimit>1</RowLimit></View>'
        $spItems = $spList.GetItems($query)
        $spCtx.Load($spItems)
        $spCtx.ExecuteQuery()
        $LastModified = $spItems[0]["Modified"]
        $OutputString = $OutputString +$delimiter+$LastModified
        Return $LastModified
    }
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_GetAllByField ($spCtx, $filterField="", $FilterValue="") {
#Could Just be All by Title or ID
    If ( ($FilterValue -ne "") -and ($filterField -eq "Title") ){
        $spLists = $spCtx.Web.Lists.GetByTitle($FilterValue)
    } Else {
		#d2. 
        $spLists = $spCtx.Web.Lists
    }
    $spCtx.Load($spLists)
    $spCtx.ExecuteQuery()

    If ( ($filterField -ne "") -and ($filterField -ne "Title") ){
        Return $spLists | Where {($_[$filterField] -eq $FilterValue)} 
    }
    Return $spLists
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_GetContentTypes ($spCtx,$spList) {
    $spContentTypes = $spList.ContentTypes
    $spCtx.Load($spContentTypes)  
	$spCtx.ExecuteQuery() 
    Return $spContentTypes
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_GetAllFields ($spCtx, $listTitle) {
	$spList = $spCtx.Web.Lists.GetByTitle($listTitle)
	$spCtx.Load($spList)
    $spFields = $spList.Fields
	$spCtx.Load($spFields)
	$spCtx.ExecuteQuery()
    Return $spFields
}
# # # # # # # # # # # # # # # # # # # #

# SET

# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_SetProp_Folders($spCtx, $listTitle, $bEnable) { 
	$spList=$spCtx.Web.Lists.GetByTitle($listTitle)
	$spList.EnableFolderCreation = $bAttachments 
	$spList.Update()
	$spCtx.ExecuteQuery()
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_SetProp_Direction($spCtx, $listTitle, $Direction="rtl") { 
	$spList=$spCtx.Web.Lists.GetByTitle($listTitle)
	$spList.Direction = $Direction 
	$spList.Update()
	$spCtx.ExecuteQuery()
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_SetProp_Direction{ 
	param (
		[Parameter(Mandatory=$true,Position=0)]
		$ctx,
		[Parameter(Mandatory=$true,Position=1)]
		[string]$ListTitle,
		[Parameter(Mandatory=$false,Position=2)]
		[string]$Title,		
		[Parameter(Mandatory=$false,Position=3)]
		[bool]$NoCrawl,
		[Parameter(Mandatory=$false,Position=4)]
		[string]$Tag,
		[Parameter(Mandatory=$false,Position=5)]
		[bool]$ContentTypesEnabled, 
		[Parameter(Mandatory=$false,Position=6)]
		[string]$Description, 
		[Parameter(Mandatory=$false,Position=7)]
		[ValidateSet(0,1,2)]
		[Int]$DraftVersionVisibility, 
		[Parameter(Mandatory=$false,Position=8)]
		[bool]$EnableMinorVersions,
		[Parameter(Mandatory=$false,Position=10)]
		[bool]$EnableAttachments,		
		[Parameter(Mandatory=$false,Position=9)]
		[bool]$EnableFolderCreation,
		[Parameter(Mandatory=$false,Position=8)]
		[bool]$EnableVersioning,
		[Parameter(Mandatory=$false,Position=8)]
		[bool]$EnableModeration,
		[Parameter(Mandatory=$false,Position=8)]
		[bool]$ForceCheckout,
		[Parameter(Mandatory=$false,Position=8)]
		[bool]$Hidden,
		[Parameter(Mandatory=$false,Position=8)]
		[bool]$IRMEnabled,
		[Parameter(Mandatory=$false,Position=8)]
		[bool]$IsApplicationList,
		[Parameter(Mandatory=$false,Position=8)]
		[bool]$OnQuickLaunch
	)
	$ll=$ctx.Web.Lists.GetByTitle($ListTitle)
	$ctx.ExecuteQuery()
	if($PSBoundParameters.ContainsKey("NoCrawl"))
		{$ll.NoCrawl=$NoCrawl}
	if($PSBoundParameters.ContainsKey("Title"))
		{$ll.Title=$Title}
	if($PSBoundParameters.ContainsKey("Tag"))
		{$ll.Tag=$Tag}
	if($PSBoundParameters.ContainsKey("ContentTypesEnabled"))
		{$ll.ContentTypesEnabled=$ContentTypesEnabled}
	if($PSBoundParameters.ContainsKey("Description"))
		{$ll.Description=$Description}
	if($PSBoundParameters.ContainsKey("DraftVersionVisibility"))
		{$ll.DraftVersionVisibility=$DraftVersionVisibility}
	if($PSBoundParameters.ContainsKey("EnableAttachments"))
		{$ll.EnableAttachments=$EnableAttachments}
	if($PSBoundParameters.ContainsKey("EnableMinorVersions"))
		{$ll.EnableMinorVersions=$EnableMinorVersions}
	if($PSBoundParameters.ContainsKey("EnableFolderCreation"))
		{$ll.EnableFolderCreation=$EnableFolderCreation}
	if($PSBoundParameters.ContainsKey("EnableVersioning"))
		{$ll.EnableMinorVersions=$EnableMinorVersions}
	if($PSBoundParameters.ContainsKey("EnableModeration"))
		{$ll.EnableModeration=$EnableModeration}
	if($PSBoundParameters.ContainsKey("ForceCheckout"))
		{$ll.ForceCheckout=$ForceCheckout}
	if($PSBoundParameters.ContainsKey("Hidden"))
		{$ll.Hidden=$Hidden}
	if($PSBoundParameters.ContainsKey("IRMEnabled"))
		{$ll.IRMEnabled=$IRMEnabled}
	if($PSBoundParameters.ContainsKey("IsApplicationList"))
		{$ll.IsApplicationList=$IsApplicationList}
	if($PSBoundParameters.ContainsKey("OnQuickLaunch"))
		{$ll.OnQuickLaunch=$OnQuickLaunch}
	$ll.Update()
	
try{
	$ctx.ExecuteQuery()
	Write-Host "Done" -ForegroundColor Green
}
catch [Net.WebException] {
	Write-Host "Failed" $_.Exception.ToString() -ForegroundColor Red
}
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_ChangeTitle ($spCtx, $spList, $listNameNew) {
    $spList.Title = $listNameNew
    $spList.Update()
    $spCtx.ExecuteQuery()
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_SetVersions ($spCtx, $ListTitle, $bMajor=$True, $bMinor=$True, $iMajorCount=$NULL, $iMinorCount=$NULL) {
	$spList = $spCtx.Web.Lists.GetByTitle($ListTitle)
	$spCtx.Load($spList)
	$ctx.ExecuteQuery()
    $spList.EnableMajorVersions = $bMajor
	$spList.
    $spList.EnableMinorVersions = $bMinor
    $spList.Update()
}
# # # # # # # # # # # # # # # # # # # #

# REMOVE

# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_RemoveField ($spCtx, $ListTitle, $FieldTitle) {
	$spList = $spCtx.Web.Lists.GetByTitle($ListTitle)
	$spCtx.Load($spList)
	$spCtx.ExecuteQuery()
	
	$spField=$spList.Fields.GetByTitle($FieldTitle)
	$spCtx.ExecuteQuery()
	$spField.DeleteObject()
	$spCtx.ExecuteQuery()
}
# # # # # # # # # # # # # # # # # # # #

# ADD

# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_AddChoiceField {
	param (
		[Parameter(Mandatory=$true,Position=1)]
		$ctx,
		[Parameter(Mandatory=$true,Position=4)]
		[string]$ListTitle,
		[Parameter(Mandatory=$true,Position=5)]
		[string]$FieldDisplayName,
		[parameter(Mandatory=$true, ValueFromPipeline=$true)]
		[String[]] $ChoiceNames,
		[Parameter(Mandatory=$false,Position=7)]
		[string]$Description="",
		[Parameter(Mandatory=$false,Position=8)]
		[string]$Required="false",
		[Parameter(Mandatory=$false,Position=9)]
		[ValidateSet('Dropdown','Radiobuttons', 'Checkboxes')]
		[string]$Format="Dropdown",
		[Parameter(Mandatory=$false,Position=10)]
		[string]$Group="",
		[Parameter(Mandatory=$true,Position=11)]
		[string]$StaticName,
		[Parameter(Mandatory=$true,Position=12)]
		[string]$Name,
		[Parameter(Mandatory=$false,Position=13)]
		[string]$Version="1"
	)

	$ctx.Load($ctx.Web)
	$ctx.Load($ctx.Web.Lists)
	$ctx.ExecuteQuery()

	$List=$ctx.Web.Lists.GetByTitle($ListTitle)
	$ctx.ExecuteQuery()

	$FieldOptions=[Microsoft.SharePoint.Client.AddFieldOptions]::AddToAllContentTypes 
	$xml="<Field Type='Choice' "
	$xml+="Description='"+$Description+"' "
	$xml+="Required='"+$Required+"' "
	$xml+="FillInChoice='FALSE'  "
	$xml+="Format='"+$Format+"' "
	$xml+="Group='"+$Group+"' "
	$xml+="StaticName='"+$StaticName+"' "
	$xml+="Name='"+$Name+"' "
	$xml+="DisplayName='"+$FieldDisplayName+"' "
	$xml+="Version='"+$Version+"' "
	$xml+="><CHOICES>"
	foreach($choice in $ChoiceNames){
		$xml+="<CHOICE>"+$choice+"</CHOICE>"
	}
	$xml+="</CHOICES></Field>"
	Write-Host $xml
	$List.Fields.AddFieldAsXml($xml,$true,$FieldOptions) 
	$List.Update() 

	try{
	$ctx.ExecuteQuery()
	Write-Host "Field " $FieldDisplayName " has been added to " $ListTitle
	}
	catch [Net.WebException]{ 
	Write-Host $_.Exception.ToString() -ForegroundColor
	}
}
# # # # # # # # # # # # # # # # # # # #
# VIEWS #
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_ViewDelete  ($ctx,$ListTitle,$ViewName) {
#https://gallery.technet.microsoft.com/scriptcenter/Remove-view-from-948763f4
    $ll=$ctx.Web.Lists.GetByTitle($ListTitle)
    $vv=$ll.Views.GetByTitle($ViewName) 
    $ctx.Load($vv)
    $ctx.ExecuteQuery()
    $vv.DeleteObject()
    $ctx.ExecuteQuery
}

# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_View_OutputAll ($spCtx,$ListTitle) {
	$ll=$spCtx.Web.Lists.GetByTitle($ListTitle)
	$spCtx.load($ll)
	$spCtx.load($ll.Views)
	$spCtx.ExecuteQuery()

	foreach($vv in $ll.Views){
		$spCtx.Load($vv)
		$spCtx.Load($vv.ViewFields)
		$spCtx.ExecuteQuery()
		Write-Output $vv
	}
}

# # # # # # # # # # # # # # # # # # # #
#DELETE ###########

# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_List_DeleteByTitle($spCtx, $listTitle) { 
 	$spList = $spCtx.Web.Lists.GetByTitle($listTitle)
    $spList.DeleteObject()
    $spCtx.ExecuteQuery
	Return "List Deleted: $listTitle"
}
