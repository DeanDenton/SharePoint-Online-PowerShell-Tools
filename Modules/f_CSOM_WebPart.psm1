#f_CSOM_WebPart.psm1

############
Function f_CSOM_WebPart_AddByXML($spCtx, $PageSvrRelURL, $WebPartTitle, $wpXML,$wpZone="Header", $wpZoneIndex="0") {
	#Can get $wpXML from exporting a WP and viewing in Notepad
	$spPage = $spCtx.Web.getFileByServerRelativeUrl($PageSvrRelURL)
	$spCtx.Load($spPage)
	$spCtx.ExecuteQuery()

	$spPage.CheckOut()
	$spCtx.ExecuteQuery()

	$wpMgr = $spPage.getLimitedWebPartManager("Shared")
	$wpDef = $wpMgr.ImportWebPart($wpxml) #D2 what is $xpxml

	$AddWP = $wpMgr.AddWebPart($wpDef.WebPart, $wpZone, $wpZoneIndex )

	$spPage.CheckIn("Initial Setup", "MajorCheckIn")
	$spPage.Publish("Initial Publish")
	$spCtx.ExecuteQuery()
}

#####################################
Function f_CSOM_WebPart_GetByName($spCtx, $wp) {


}


Function f_CSOM_WebPart_ReturnAll($spCtx, $Recursive=$true){
	$spCtx.Load($spCtx.Web)
	$spCtx.ExecuteQuery()

	$spqQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
	# $spqQuery.ViewAttributes = "Scope='Recursive'"

	if($Recursive)	{		$spqQuery.ViewXml ="<View Scope='RecursiveAll' />"	}

	$itemki = $spCtx.Web.Lists.GetbyTitle("Site Pages").GetItems($spqQuery)
	$spCtx.Load($itemki)
	$spCtx.ExecuteQuery()

	foreach($item in $itemki){
Write-Host "*****" $item.Client_Title
		$file = $item.File
		$spCtx.Load($file)
		$spCtx.ExecuteQuery()

		$page = $spCtx.Web.GetFileByServerRelativeUrl($file.ServerRelativeUrl)
		$wpm = $page.GetLimitedWebPartManager("Shared")
		$spCtx.Load($wpm);
		$spCtx.Load($wpm.WebParts);
		$spCtx.ExecuteQuery()

		foreach($webbie in $wpm.WebParts)	{
			$obj= New-Object PSObject
			$obj | Add-Member NoteProperty -Name Page -Value $file.ServerRelativeUrl 
			$obj | Add-Member NoteProperty -Name DefinitionID -Value $webbie.ID
			$spCtx.Load($webbie.WebPart)
			$spCtx.Load($webbie.WebPart.Properties)
			$spCtx.ExecuteQuery()
			
			$obj | Add-Member NoteProperty -Name IsClosed -Value $webbie.WebPart.IsClosed
			$obj | Add-Member NoteProperty -Name Hidden -Value $webbie.WebPart.Hidden
			$obj | Add-Member NoteProperty -Name Subtitle -Value $webbie.WebPart.Subtitle
			$obj | Add-Member NoteProperty -Name Title -Value $webbie.WebPart.Title
			$obj | Add-Member NoteProperty -Name TitleUrl -Value $webbie.WebPart.TitleUrl
			$obj | Add-Member NoteProperty -Name ZoneIndex -Value $webbie.WebPart.ZoneIndex
			$obj | Add-Member NoteProperty -Name ServerObjectIsNull -Value $webbie.WebPart.ServerObjectIsNull

			foreach($fv in $webbie.WebPart.Properties.FieldValues){
				$obj | Add-Member NoteProperty -Name $fv -Value $fv
			}
Write-Output $obj
Write-Host "-------------------------------------------------------------------------------------------"-BackgroundColor Cyan
			#Export-Csv -InputObject $obj -Append -LiteralPath C:\test634.csv
		}
        
		$page=$null
	}
}



############
#Doesn't Work Yet
Function f_CSOM_WebPart_DeleteAllOnPage($spCtx, $PageSvrRelURL) {
	$spPage = $spCtx.Web.GetFileByServerRelativeUrl($PageSvrRelURL)
	$wpMgr = $spPage.GetLimitedWebPartManager("Shared")
	$spCtx.Load($wpMgr);
	$spCtx.Load($wpMgr.WebParts);
	$spCtx.ExecuteQuery()
	
	$wpMgr.WebParts.Count
	foreach($spWpmPart in $wpMgr.WebParts){
        $spWebPart = $spWpmPart.WebPart
        $spWpProperties = $spWpmPart.WebPart.Properties
	    $spCtx.Load($spWebPart)
	    $spCtx.Load($spWpProperties)
	    $spCtx.ExecuteQuery()
        $spWebPart  | Out-GridView 
        #$spWpProperties| Out-GridView
		$spWpProperties.fieldvaluse | Out-GridView
        #$spWebPart.DeleteWebPart()
	}
	$spCtx.ExecuteQuery()
}


#####################################
Function f_CSOM_WebPart_GetAll{
#https://gallery.technet.microsoft.com/scriptcenter/Create-a-report-on-all-web-55c57a47
	param (
		[Parameter(Mandatory=$true,Position=1)]
		$ctx,
		[Parameter(Mandatory=$true,Position=2)]
		[bool]$Recursive
	)
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()

	$spqQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
	# $spqQuery.ViewAttributes = "Scope='Recursive'"

	if($Recursive)	{
		$spqQuery.ViewXml ="<View Scope='RecursiveAll' />";
	}

	$itemki=$ctx.Web.Lists.GetbyTitle("Site Pages").GetItems($spqQuery)
	$ctx.Load($itemki)
	$ctx.ExecuteQuery()
	foreach($item in $itemki){
		Write-Host "*****" $item.Client_Title
		$file = $item.File
		$ctx.Load($file)
		$ctx.ExecuteQuery()
		$page = $ctx.Web.GetFileByServerRelativeUrl($file.ServerRelativeUrl)
		$wpm = $page.GetLimitedWebPartManager("Shared")
		$ctx.Load($wpm);
		$ctx.Load($wpm.WebParts);
		$ctx.ExecuteQuery()

		foreach($webbie in $wpm.WebParts)	{
			$obj= New-Object PSObject
			$obj | Add-Member NoteProperty -Name Page -Value $file.ServerRelativeUrl 
			$obj | Add-Member NoteProperty -Name DefinitionID -Value $webbie.ID
			$ctx.Load($webbie.WebPart)
			$ctx.Load($webbie.WebPart.Properties)
			$ctx.ExecuteQuery()
			
			$obj | Add-Member NoteProperty -Name IsClosed -Value $webbie.WebPart.IsClosed
			$obj | Add-Member NoteProperty -Name Hidden -Value $webbie.WebPart.Hidden
			$obj | Add-Member NoteProperty -Name Subtitle -Value $webbie.WebPart.Subtitle
			$obj | Add-Member NoteProperty -Name Title -Value $webbie.WebPart.Title
			$obj | Add-Member NoteProperty -Name TitleUrl -Value $webbie.WebPart.TitleUrl
			$obj | Add-Member NoteProperty -Name ZoneIndex -Value $webbie.WebPart.ZoneIndex
			$obj | Add-Member NoteProperty -Name ServerObjectIsNull -Value $webbie.WebPart.ServerObjectIsNull

			foreach($fv in $webbie.WebPart.Properties.FieldValues){
				$obj | Add-Member NoteProperty -Name $fv -Value $fv
			}
			Write-Output $obj
			Write-Host "-------------------------------------------------------------------------------------------"-BackgroundColor Cyan
			#Export-Csv -InputObject $obj -Append -LiteralPath C:\test634.csv
		}
		$page=$null
	}
}
#####################################
Function f_CSOM_WebPart_GetAllPropertiesPerPage{
	param (
        [Parameter(Mandatory=$true,Position=1)]
		$ctx,
		[Parameter(Mandatory=$true,Position=2)]
		[string]$pageUrl
	)
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()

	$page = $ctx.Web.GetFileByServerRelativeUrl($pageUrl)
	$wpm = $page.GetLimitedWebPartManager("Shared")
	$ctx.Load($wpm);
	$ctx.Load($wpm.WebParts);
	$ctx.ExecuteQuery()

	foreach($webbie in $wpm.WebParts){
		Write-Output $webbie
		$ctx.Load($webbie.WebPart)
		$ctx.Load($webbie.WebPart.Properties)
		$ctx.ExecuteQuery()
		Write-Host "Associated web part:" -ForegroundColor DarkGreen
		Write-Output $webbie.WebPart
		Write-Output $webbie.WebPart.Properties.FieldValues 
		Write-Host "--------------------------------------------------" -BackgroundColor Cyan
	}
}
############
Function f_CSOM_WebPart_DeleteOnePerPage {
#https://gallery.technet.microsoft.com/scriptcenter/Delete-single-web-part-6404fced
	param (
        [Parameter(Mandatory=$true,Position=1)]
		$ctx,
		[Parameter(Mandatory=$true,Position=2)]
		[string]$pageUrl,
		[Parameter(Mandatory=$true,Position=5)]
		[string]$webPartID
	)
	
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()

	$page = $ctx.Web.GetFileByServerRelativeUrl($pageUrl)
	$wpm = $page.GetLimitedWebPartManager("Shared")
	$ctx.Load($wpm);
	$ctx.Load($wpm.WebParts);
	$ctx.ExecuteQuery()
	
	$webbie=$wpm.WebParts.GetById($webPartID)
	Write-Host "Deleting web part id: " $webPartID
	$webbie.DeleteWebPart()
	$ctx.ExecuteQuery()
}

############
function f_CSOM_WebPart_GetByID {}
############
 
function f_CSOM_WebPart_DeleteAllPerPage {
#https://gallery.technet.microsoft.com/scriptcenter/Delete-all-web-parts-from-4c7bc264
	param (
        [Parameter(Mandatory=$true,Position=1)]
		$ctx,
		[Parameter(Mandatory=$true,Position=2)]
		[string]$pageUrl
	)
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()

	$page = $ctx.Web.GetFileByServerRelativeUrl($pageUrl)
	$wpm = $page.GetLimitedWebPartManager("Shared")
	$ctx.Load($wpm);
	$ctx.Load($wpm.WebParts);
	$ctx.ExecuteQuery()

	foreach($webbie in $wpm.WebParts)            {
		Write-Host "Deleting web part id: " $webbie.Id
		$webbie.DeleteWebPart()
		$ctx.ExecuteQuery()
	}
}
