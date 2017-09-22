#cls
######################################
$ReuseCredentials = $true
If (($Credentials -eq $Null) -or ($ReuseCredentials -eq $false)) { $Credentials = Get-Credential }

$isSPO = $true  
 # CONTROL #
$ControlURL = "https://maritzllc.sharepoint.com/SUP/ws"
$dataWebsTitle = "dataWebs"
$dataWebAppsTitle = "dataWebApps"

$SourceListTitle="dataSiteCols"
$GroupListTitle="dataSiteColGroups"

$filterFlag = "Scan"

#$updateALL = $false
$ScanDateTime = Get-Date
######################################
# ADD: Navigation Info

######################################
$CSOM_Path   = "C:\DEV\Microsoft.SharePointOnline.CSOM.16.1.6906.1200\lib\net45\"
$ModulesPath = "C:\DEV\GitHub\SharePoint-Online-PowerShell-Tools\Modules\"
######################################
 # IMPORT #
Add-Type -Path "$CSOM_Path\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "$CSOM_Path\Microsoft.SharePoint.Client.Runtime.dll" 
######################################
 # MODULES #
import-Module -Name ($ModulesPath + "f_CSOM_Ctx.psm1") -force
import-Module -Name ($ModulesPath + "f_CSOM_List.psm1") -force
import-Module -Name ($ModulesPath + "f_CSOM_Item.psm1") -force
######################################

Function f_CSOM_Web_UpdateFields ($webURL) {
    $color = "green"  

    $IsDataWebApp = $webURL.StartsWith("https://maritzllc.sharepoint.com")

    If ($IsDataWebApp) {
        $ControlListTitle = $dataWebsTitle
    }
    Else{
        $ControlListTitle = $dataWebAppsTitle
    }


    $itemTitle = $webURL.Replace("https://maritzllc.sharepoint.com","")
    $queryXML = '<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+ $itemTitle+ '</Value></Eq></Where></Query></View>'
    $spControlItem = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Control -listTitle $ControlListTitle -queryXML $queryXML 

    If ($spControlItem.Count -eq 0) {
        $spControlItem = f_CSOM_Item_Add -spCtx $spCtx_Control -listTitle $ControlListTitle  -itemTitle $itemTitle
        Write-Host "New URL: " $webUrl  -foregroundcolor DarkYellow
    } 

    $spCtxWeb = f_CSOM_Ctx_Get -CtxURL $webURL -Credentials $Credentials -isSPO $isSPO 
    $spWeb = $spCtxWeb.Web 
    $spCtxWeb.Load($spWeb) 
    $spCtxWeb.ExecuteQuery() 

    Write-Host "Update Item: " $itemTitle " ("$spWeb.Title")" -foregroundcolor $color
    $queryXML = '<View><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+ $itemTitle+ '</Value></Eq></Where><RowLimit>1</RowLimit></View>'
    $spItem = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Control -listTitle $ControlListTitle -queryXML $queryXML

<#        
        $spCtx.Load($spWeb.Features)
        ("Feature Count: " + $spWeb.Features.Count) 
#>  
     $spCtxWeb.Load($spWeb.RootFolder)
     #$spCtxWeb.Load($spWeb.RegionalSettings.TimeZone)
     $spCtxWeb.ExecuteQuery()

     Try {
        $spCtxWeb.Load($spWeb.AssociatedOwnerGroup)
        $spCtxWeb.Load($spWeb.AssociatedOwnerGroup.Users)
        $spCtxWeb.Load($spWeb.AssociatedMemberGroup)
        $spCtxWeb.Load($spWeb.AssociatedVisitorGroup)
        $spCtxWeb.Load($spWeb.RegionalSettings.TimeZone)
        $spCtxWeb.ExecuteQuery()
        $OwnerGroupTitle = $spWeb.AssociatedOwnerGroup.Title
        $OwnerGroupOwner = $spWeb.AssociatedOwnerGroup.OwnerTitle
        $OwnerGroupUsers = ( $spWeb.AssociatedOwnerGroup.Users.Title -join "; " )
        $MemberGroupTitle = $spWeb.AssociatedMemberGroup.Title
        $MemberGroupOwner = $spWeb.AssociatedMemberGroup.OwnerTitle
        $VisitorGroupTitle = $spWeb.AssociatedVisitorGroup.Title
        $VisitorGroupOwner = $spWeb.AssociatedVisitorGroup.OwnerTitle
        $TimeZone = $spWeb.RegionalSettings.TimeZone.Description
    } 
    Catch {
        $OwnerGroupTitle = "error"
        $OwnerGroupOwner = "error"
        $OwnerGroupUsers = "error"
        $MemberGroupTitle = "error"
        $MemberGroupOwner = "error"
        $VisitorGroupTitle = "error"
        $VisitorGroupOwner = "error"
        $TimeZone = "error"
    }

        Connect-PnPOnline -Url $webURL -Credentials $Credentials
        $AccessRequest = Get-PnPRequestAccessEmails 
        $GetPnPTheme = Get-PnPTheme
  
        $WebFeatures = ""
        Get-PnPFeature | Sort-Object DisplayName | %{ $WebFeatures = $WebFeatures + ", " + $_.DisplayName } 
        If ($WebFeatures.Length -gt 0) {$WebFeatures = $WebFeatures.Substring(2,$WebFeatures.Length-2)}

    $dataHASH = @{
        "WebLastScan" = $ScanDateTime.ToUniversalTime() 
        "WebLastModified" = $spWeb.LastItemModifiedDate
        "WebTemplate" = $spWeb.WebTemplate
        "ListCount" = f_CSOM_List_Count -spCtx $spCtxWeb -visibleOnly $true
        "ListItemCountMAX" = f_CSOM_List_MaxItemCount -spCtx $spCtxWeb -visibleOnly $true
        "ListItemCount" = f_CSOM_Lists_FullItemCount -spCtx $spCtxWeb -visibleOnly $true
        "MasterUrl" = $spWeb.MasterUrl  
        "CustomMasterUrl" = $spWeb.CustomMasterUrl  
        "AlternateCssUrl" = $spWeb.AlternateCssUrl
        "SiteLogoUrl" = $spWeb.SiteLogoUrl
        "WebTitle" = $spWeb.Title
        "WebDescription" = $spWeb.Description
        "WebUrl" = $webURL
        "SiteColLookup" = $SiteColID
        "OwnerGroupTitle" = $OwnerGroupTitle
        "OwnerGroupOwner" = $OwnerGroupOwner 
        "OwnerGroupUsers" = $OwnerGroupUsers
        "MemberGroupTitle" = $MemberGroupTitle
        "MemberGroupOwner" = $MemberGroupOwner
        "VisitorGroupTitle" = $VisitorGroupTitle
        "VisitorGroupOwner" = $VisitorGroupOwner
        "WelcomePage" = $spWeb.RootFolder.WelcomePage
        "AccessRequest" = $AccessRequest
        "Theme" =  $GetPnPTheme.Name
        "ThemePath" =  $GetPnPTheme.Theme
        "ThemeCustom" = $GetPnPTheme.IsCustomComposedLook
        "ThemeBackground" = $GetPnPTheme.BackgroundImage
        "ThemeFont" = $GetPnPTheme.Font
        "Features" =  $WebFeatures
        "TimeZone" = $TimeZone
    }
    #$dataHASH | Out-GridView

    f_CSOM_Item_UpdateByTitle -spCtx $spCtx_Control -ListTitle $ControlListTitle -ItemTitle $itemTitle -spItemFieldsHash $dataHASH

    #AllowCreateDeclarativeWorkflow, AllowDesigner, Owner, PrimaryUri, 
}

Function f_CSOM_SiteCol_GetUsers ($spCtx_Control, $siteURL)  {
    $SiteTitle = $siteURL.Replace("https://maritzllc.sharepoint.com","")

    $spCtx_Site = f_CSOM_Ctx_Get -CtxURL $siteURL -Credentials $Credentials -getCredAlways $false -isSPO $true
    $spRootWebSite = $spCtx_Site.Web 
    $spSites = $spRootWebSite.Webs 

    $spGroups=$spCtx_Site.Web.SiteGroups 
    $spCtx_Site.Load($spGroups) 
    $spCtx_Site.ExecuteQuery()        
         
    foreach($spGroup in $spGroups){ 
        $spCtx_Site.Load($spGroup) 
        $spCtx_Site.ExecuteQuery() 
        Write-Host "* " $spGroup.Title 
 
 <#
        #Getting the users per group in the SPO Site 
        $spSiteUsers=$spGroup.Users 
        $spCtx_Site.Load($spSiteUsers) 
        $spCtx_Site.ExecuteQuery() 
        foreach($spUser in $spSiteUsers){ 
            Write-Host "    -> User:" $spUser.Title " - User ID:" $spUser.Id " - User E-Mail" $spUser.Email " - User Login" $spUser.LoginName                 
        } 
#>

        $itemTitle = $SiteTitle + "|" + $spGroup.Title 
        $queryXML = '<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+ $itemTitle+ '</Value></Eq></Where></Query></View>'
        $spGroupItem = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Control -listTitle $GroupListTitle -queryXML $queryXML 

        If ($spGroupItem.Count -eq 0) {
            $spGroupItem = f_CSOM_Item_Add -spCtx $spCtx_Control -listTitle $GroupListTitle  -itemTitle $itemTitle
            Write-Host "New Group: " $itemTitle  -foregroundcolor Green
        } 

        Write-Host "Update Item: " $itemTitle " ("$spWeb.Title")" -foregroundcolor Blue
        $queryXML = '<View><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+ $itemTitle+ '</Value></Eq></Where><RowLimit>1</RowLimit></View>'
        $spItem = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Control -listTitle $ControlListTitle -queryXML $queryXML

        $dataHASH = @{
        "SiteColLookup" = $SiteColID
        }
        #f_CSOM_Item_UpdateByTitle -spCtx $spCtx_Control -ListTitle $GroupListTitle -ItemTitle $itemTitle -spItemFieldsHash $dataHASH
    } 
}

#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-# 
Function f_CSOM_Web_GetSubSites($webUrl, $recurs) {
    $spCtx = f_CSOM_Ctx_Get -CtxURL $webURL -Credentials $Credentials -isSPO $isSPO                   
    $spWeb = $spCtx.Web            
    $spCtx.Load($spWeb) 
    f_CSOM_Web_UpdateFields -webURL $webURL 

    If ($recurs -eq $true) {
        $spSites = $spWeb.Webs
        $spCtx.Load($spSites)
        $spCtx.ExecuteQuery()

        foreach($spWeb in $spSites) {
            f_CSOM_Web_GetSubSites -webUrl $spWeb.Url -recurs $recurs
        }
    }
}
 
#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-# 

######################################
 # MAIN #

# GET CONTEXT # # # # # # # # # # # #  
$spCtx_Control = f_CSOM_Ctx_Get -CtxURL $ControlURL -Credentials $Credentials -isSPO $isSPO


# GET CONTROL LIST ITEMS # # # # # # # # # # # #  
<#
$queryXML = '<View><Query><Where><And>`
    <Eq><FieldRef Name="Site_x0020_Type"/><Value Type="Text">Site</Value></Eq>`
    <Eq><FieldRef Name="SiteDir"/><Value Type="Text">Show</Value></Eq>`
    </And></Where></Query></View>'
#>

$queryXML = '<View><Query><Where><Eq><FieldRef Name="Flag"/><Value Type="Text">'+$filterFlag+'</Value></eq></Where></Query></View>'
$spSiteColsItems = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Control -listTitle $SourceListTitle -queryXML $queryXML

$ScanDateTime
 ForEach ($spItem in $spSiteColsItems) {
    #$spItem| Out-GridView
    $SiteColURL = $spItem["SiteColURL"]
    $SiteColID =  $spItem["ID"]
    #f_CSOM_SiteCol_GetUsers -spCtx_Control $spCtx_Control -siteURL $SiteColURL 
    f_CSOM_Web_GetSubSites -webURL $SiteColURL -recurs $true
}
