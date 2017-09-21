#f_CSOM_Item.psm1

## GET ############################### 

 Function f_CSOM_Item_GetByID ($spCtx, $listTitle, $itemID) {
    $spList = $spCtx.web.Lists.GetByTitle($listTitle)
    $spCtx.ExecuteQuery()
	
    $spItem = $spList.GetItemById($itemID)
    $spCtx.Load($spItem)
    $spCtx.ExecuteQuery()
}
# # # # # # # # # # # # # # # # # # # #
 Function f_CSOM_Item_GetAllByQuery ($spCtx, $listTitle, $queryXML="") {
	$spList = $spCtx.Web.Lists.GetByTitle($listTitle)
	$spCtx.Load($spList)
	$spCtx.ExecuteQuery()
	$query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXml = $queryXML
    $spItems = $spList.GetItems($query)
    $spCtx.Load($spItems)
    $spCtx.ExecuteQuery()
	Return $spItems 
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Item_GetAllByItemTitle ($spCtx, $listTitle, $itemTitle) {
    $spList = $spCtx.web.Lists.GetByTitle($listTitle)
    $spCtx.ExecuteQuery()
    $queryXML = '<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+$ItemTitle+'</Value></Eq></Where></Query></View>'
    $spItems = f_CSOM_Item_GetAllByQuery $spCtx $listTitle $queryXML  
    Return $spItems
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Item_GetAllByListTitle ($spCtx, $listTitle, $queryXML="") {
    $queryXML = '<View Scope="RecursiveAll"></View>'
    $spItems = f_CSOM_Item_GetAllByQuery -spCtx $spCtx -listTitle $listTitle -queryXML $queryXML  
	Return $spItems 
}

######################################
 # UPDATE #
# # # # # # # # # # # # # # # # # # # #

Function f_CSOM_Item_UpdateByTitle ($spCtx, $ListTitle, $ItemTitle, $spItemFieldsHash) { 
    <# HAsH EXAMPLE
    $spItemFieldsHash = @{
        "<FieldTitle1>" = <FieldValue1>
        "<FieldTitleN>" = <FieldValueN>
    }
    #>
	$spCtx.Load($spCtx.Web)
	$spCtx.Load($spCtx.Web.Lists)
	$spCtx.ExecuteQuery()
    $queryXML = '<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+$ItemTitle+'</Value></Eq></Where></Query></View>'

    $spItems = f_CSOM_Item_GetAllByQuery $spCtx $listTitle $queryXML  
    ForEach ($spItem in $spItems) {
	#$spItem.FieldValues
         ForEach ($Key in $spItemFieldsHash.Keys){
             $spItem[$Key] = $spItemFieldsHash[$Key]
        } 
        $spItem.Update() 
    }
    Try {
        $spCtx.ExecuteQuery() 
    }
    Catch {}
}

######################################
 # ADD #
Function f_CSOM_Item_Add($spCtx, $listTitle, $itemTitle, $spItemFieldsHash="") {
	$spCtx.Load($spCtx.Web)
	$spCtx.Load($spCtx.Web.Lists)
	$spCtx.ExecuteQuery()

	$spList=$spCtx.Web.Lists.GetByTitle($ListTitle)
	$spCtx.Load($spList)
	$spCtx.ExecuteQuery()

    $spItemCreateInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
    $spNewItem = $spList.AddItem($spItemCreateInfo)
    $spNewItem["Title"] = $itemTitle
    ForEach ($Key in $spItemFieldsHash.Keys){
        $spNewItem[$Key] = $spItemFieldsHash[$Key]
    } 
    $spNewItem.Update() 
    $spCtx.ExecuteQuery() 
    Return $spNewItem
}

#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-# 
Function f_CSOM_Item_AddNewTitle($spCtx, $ListTitle, $ItemTitle) {
    $queryXML = '<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+ $ItemTitle+ '</Value></Eq></Where></Query></View>'
    $spItems = f_CSOM_Item_GetAllByListTitle -spCtx $spCtx -listTitle $ListTitle -queryXML $queryXML 

    If ($spItems.Count -eq 0) {
        $spItems = f_CSOM_Item_Add -spCtx $spCtx -listTitle $ListTitle -itemTitle $ItemTitle
        #Write-Host "New URL: " $ItemTitle  -foregroundcolor Green
    } 
}

# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Item_UpdateTitle ($ctx, $ListTitle, $ItemTitle, $OtherFieldName="", $OtherFieldValue="") { 
	$ctx.Load($ctx.Web)
	$ctx.Load($ctx.Web.Lists)
	$ctx.ExecuteQuery()

	$ll=$ctx.Web.Lists.GetByTitle($ListTitle)
	$ctx.Load($ll)
	$ctx.ExecuteQuery()

	$lici =New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
	$listItem = $ll.AddItem($lici)
	$listItem["Title"]=$ItemTitle
	if($OtherFieldName -ne ""){
		$listItem[$OtherFieldName]=$OtherFieldValue
	}
	$listItem.Update()
	$ll.Update()
	$ctx.ExecuteQuery()
}
#####################################
 # DELETE #
Function f_CSOM_Item_DeleteByID ($spCtx, $listTitle, $itemID) {
    $spList = $spCtx.web.Lists.GetByTitle($listTitle)
    $spCtx.ExecuteQuery()
	
    $spItem = $spList.GetItemById($itemID)
    $spCtx.Load($spItem)
    $spCtx.ExecuteQuery()
	
    $spItem.DeleteObject()
    $spCtx.ExecuteQuery()
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Item_DeleteByQuery($spCtx, $listTitle, $queryXML) {
    $spItems = f_CSOM_Item_GetAllByQuery $spCtx $listTitle $queryXML 
    if ($spItems.Count -gt 0) {
        forEach ($spItem in $spItems) {
           $spItem["Title"]
           $spItem.DeleteObject()
        }
        $spCtx.ExecuteQuery()
    }
}

# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Item_RecycleBinAll ($spCtx, $Action="Report" ) {
#https://gallery.technet.microsoft.com/scriptcenter/Create-a-report-on-deleted-496ce018
	$spCtx.Load($spCtx.Site)
	$rb=$spCtx.Site.RecycleBin
	$spCtx.Load($rb)
	$spCtx.ExecuteQuery()
	for($i=0;$i -lt $rb.Count ;$i++){
		$obj = $rb[$i]
		Write-Output $obj
	}
	If ($Action -eq "Restore") {
		$spCtx.Web.RecycleBin.RestoreAll()
	}
	$spCtx.ExecuteQuery()
} 
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Item_RecycleBinByUser ($ctx, $UserUpn, $Action="Report") {
#Consider Merging with f_CSOM_Item_RecycleBinAll
#https://gallery.technet.microsoft.com/scriptcenter/Naughty-Employee-restore-638fcdb3
	$myarray=@()
	$ctx.Load($ctx.Site)
	$ctx.Load($ctx.Web.Webs)
	$rb=$ctx.Site.RecycleBin
	$ctx.Load($rb)
	$ctx.ExecuteQuery()

	for($i=0;$i -lt $rb.Count ;$i++)	{
		$ctx.Load($rb[$i].Author)
		$ctx.Load($rb[$i].DeletedBy)
		$ctx.ExecuteQuery()
		$obj = $rb[$i]
		$obj | Add-Member NoteProperty AuthorLoginName($rb[$i].Author.LoginName)
		$obj | Add-Member NoteProperty DeletedByLoginName($rb[$i].DeletedBy.LoginName)
		$myarray+=$obj
	}
	#Write-Output $obj
	for($i=0;$i -lt $myarray.Count ; $i++){
		if($myarray[$i].DeletedByLoginName -eq $UserUpn ){

			If ($Action -eq "Restore") {
				$myarray[$i].Restore()
			} Else {
				Write-Host "Item($i): " $myarray[$i].Title
			}
			$ctx.ExecuteQuery()
		}
	}
} 
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Item_VersionReportOnList ($ctx, $ListTitle,$CSVPath,$CSVPath2) {
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()

	$ll=$ctx.Web.Lists.GetByTitle($ListTitle)
	$ctx.Load($ll)
	$ctx.ExecuteQuery()

	$spqQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
	$spqQuery.ViewXml ="<View Scope='RecursiveAll' />";
	$itemki=$ll.GetItems($spqQuery)
	$ctx.Load($itemki)
	$ctx.ExecuteQuery()

	foreach($item in $itemki)  {
		Write-Host $item["FileRef"]
		$file = $ctx.Web.GetFileByServerRelativeUrl($item["FileRef"]);
		$ctx.Load($file)
		$ctx.Load($file.Versions)
		$ctx.ExecuteQuery()
		if ($file.Versions.Count -eq 0){
			$obj=New-Object PSObject
			$obj | Add-Member NoteProperty ServerRelativeUrl($file.ServerRelativeUrl)
			$obj | Add-Member NoteProperty FileLeafRef($item["FileLeafRef"])
			$obj | Add-Member NoteProperty Versions("No Versions Available")
			$obj | export-csv -Path $CSVPath2 -Append
		}

		Foreach ($versi in $file.Versions){
			$user=$versi.CreatedBy
			$ctx.Load($versi)
			$ctx.Load($user)
			$ctx.ExecuteQuery()
			$versi | Add-Member NoteProperty CreatedByUser($user.LoginName)
			$versi | Add-Member NoteProperty FileLeafRef($item["FileLeafRef"])
			$versi |export-csv -Path $CSVPath -Append
		}
	}
}

<#
for($j=0;$j -lt $itemki.Count ;$j++)  {
     $itemki[$j].ResetRoleInheritance()
  }
   $spCtx.ExecuteQuery()
#>