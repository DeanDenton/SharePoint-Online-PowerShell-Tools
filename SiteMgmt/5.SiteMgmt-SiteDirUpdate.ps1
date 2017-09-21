
######################################
#$CSOM_Path = "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\"
#$ModulesPath = "D:\TOOLS\Scripts\Modules\"
######################################
$CSOM_Path = "C:\SYNC\SharePoint\MSH Development - Code\dllCSOM\"
$ModulesPath = "C:\SYNC\SharePoint\MSH Development - Code\Modules\"
######################################
$ReuseCredentials = $true 
# GET CREDENTIALS  # # # # # # # # # #  
If (($Credentials -eq $Null) -or ($ReuseCredentials -eq $false)) { $Credentials = Get-Credential }

$isSPO = $true
$updateALL = $false
$ScanDateTime = Get-Date
######################################
 # IMPORT #
#$CSOM_Path = "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI"
Add-Type -Path ($CSOM_Path + "Microsoft.SharePoint.Client.dll" )
Add-Type -Path ($CSOM_Path + "Microsoft.SharePoint.Client.Runtime.dll")

######################################
 # MODULES #
import-Module -Name ($ModulesPath + "f_CSOM_Ctx.psm1") -force
import-Module -Name ($ModulesPath + "f_CSOM_List.psm1") -force
import-Module -Name ($ModulesPath + "f_CSOM_Item.psm1") -force
import-Module -Name ($ModulesPath + "f_CSOM_Perm.psm1") -force
######################################

## FUNCTIONS #########################

Function f_SiteMgmt_SiteDir_Update() {
    #Update PWA.Site Migration Task from projSP13Mig.dataWebs
    $SrcURL = "https://maritzllc.sharepoint.com/SUP/WS/"
    $SrcListTitle = "Site Management"
    $queryXML = '<View><Query><Where><Eq><FieldRef Name="OnSiteDir"/><Value Type="Text">SiteDirOn</Value></Eq></Where></Query></View>'

    $spCtx_Src = f_CSOM_Ctx_Get -CtxURL $SrcURL -Credentials $Credentials -isSPO $isSPO
    $spItems_Src = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Src -listTitle $SrcListTitle -queryXML $queryXML

    $DestURL = "https://maritzllc.sharepoint.com/"
    $DestListTitle = "Site Directory"
    $spCtx_Dest = f_CSOM_Ctx_Get -CtxURL $DestURL -Credentials $Credentials -isSPO $isSPO

    ForEach ($spItem_Src in $spItems_Src) {
        $itemTitle = $spItem_Src["Title"]

        $queryXML = '<View><Query><Where><Eq><FieldRef Name="ID"/><Value Type="ID">'+$spItem_Src["WebData"].LookupID+'</Value></Eq></Where></Query></View>'
        $spItems_Web = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Src -listTitle "dataWebs" -queryXML $queryXML
        #$spItems_Web | Out-GridView

        $webTitle = $spItems_Web["WebTitle"]          
        $WebUrl = $spItems_Web["WebUrl"] 
        $WebDescription = $spItems_Web["WebDescription"] 

        $DataHash = @{
            "Title" = $itemTitle
            "SiteTitle" = $webTitle
            "SiteLink" = $WebUrl
            "Description" = $WebDescription
            "SiteOwner" = $spItem_Src["Site_x0020_Owner"]
            "BackupContact" = $spItem_Src["Site_x0020_Manager"]
            "SitePortal" = $spItem_Src["Site_x0020_Portal"]
            "SiteClass" = $spItem_Src["SiteClass"]
            "SiteType" = $spItem_Src["Site_x0020_Type"]
            "BU" = $spItem_Src["BU"]
        }
        Write-Host "Sync: " $itemTitle 

        $queryXML = '<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+$itemTitle+'</Value></Eq></Where></Query></View>'
        $spItem_Dest = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Dest -listTitle $DestListTitle -queryXML $queryXML 
     
        If ($spItem_Dest.Count -eq 0) {
            $spControlItem = f_CSOM_Item_Add -spCtx $spCtx_Dest -listTitle $DestListTitle -itemTitle $itemTitle
            Write-Host "New itemTitle: " $itemTitle  -foregroundcolor Green
        } 
       f_CSOM_Item_UpdateByTitle -spCtx $spCtx_Dest -ListTitle $DestListTitle -ItemTitle $itemTitle -spItemFieldsHash $DataHash
    }
}

#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-# 
Function f_SigteMgmt_SiteDir_Delete() {
    #Update PWA.Site Migration Task from projSP13Mig.dataWebs
    $SrcURL = "https://maritzllc.sharepoint.com/SUP/WS/"
    $SrcListTitle = "Site Management"
    $queryXML = '<View><Query><Where><Eq><FieldRef Name="OnSiteDir"/><Value Type="Text">0</Value></Eq></Where></Query></View>'

    $spCtx_Src = f_CSOM_Ctx_Get -CtxURL $SrcURL -Credentials $Credentials -isSPO $isSPO
    $spItems_Src = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Src -listTitle $SrcListTitle -queryXML $queryXML

    $DestURL = "https://maritzllc.sharepoint.com/"
    $DestListTitle = "Site Directory"
    $spCtx_Dest = f_CSOM_Ctx_Get -CtxURL $DestURL -Credentials $Credentials -isSPO $isSPO

    ForEach ($spItem_Src in $spItems_Src) {
        $itemTitle = $spItem_Src["Title"]
        $DataHash = @{
            "OnSiteDir" = "OnSiteDirOff"
        }
       f_CSOM_Item_UpdateByTitle -spCtx $spCtx_Src -ListTitle $SrcListTitle -ItemTitle $itemTitle -spItemFieldsHash $DataHash

       $queryXML = '<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+$itemTitle+'</Value></Eq></Where></Query></View>'
       Write-Host "Delete " $itemTitle
       f_CSOM_Item_DeleteByQuery -spCtx $spCtx_Dest -listTitle $DestListTitle -queryXML $queryXML
    }
}

## MAIN ##############################

f_SiteMgmt_SiteDir_Update

#f_SigteMgmt_SiteDir_Delete