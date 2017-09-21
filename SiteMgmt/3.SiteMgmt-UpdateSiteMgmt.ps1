#cls
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

#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-# 
Function f_SiteMgmt_dataWebs_Update() {
    #Update PWA.Site Migration Task from projSP13Mig.dataWebs
    $SrcURL = "https://maritzllc.sharepoint.com/SUP/WS/"
    #$Title_Src = "/teams/MMS/MMSTech/InfoSec"
    $ListTitle_Src="dataWebs"
    $spCtx_Src = f_CSOM_Ctx_Get -CtxURL $SrcURL -Credentials $Credentials -isSPO $isSPO

    #Filter just those items from the most recent scan

    #All from last scan
    <#
    [DateTime]$MaxWebLastScan = f_CSOM_List_GetMaxField -spCtx $spCtx_Src -listName $ListTitle_Src -fieldName "WebLastScan"
    $MaxWebLastScanDate = $MaxWebLastScan.ToString("yyyy-MM-dd-hh-mm")
    $queryXML = '<View><Query><Where><Eq><FieldRef Name="WebLastScan"/><Value Type="DateTime">'+ $MaxWebLastScanDate +'</Value></Eq></Where></Query></View>'
    #>

    #All by Flag
    $Flag = "SiteMgmtOn"
    If ($Flag -ne "") {
        $queryXML = '<View><Query><Where><Eq><FieldRef Name="Flag"/><Value Type="Text">'+ $Flag +'</Value></Eq></Where></Query></View>'
    }
    Else {
        $queryXML = ""
    }

    $spItems_Src = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Src -listTitle $ListTitle_Src -queryXML $queryXML
    #$spItems_Src | out-gridview

    $DestURL = "https://maritzllc.sharepoint.com/SUP/WS/"
    $spCtx_Dest = f_CSOM_Ctx_Get -CtxURL $DestURL -Credentials $Credentials -isSPO $isSPO
    $ListTitle_Dest="Site Management"

    ForEach ($spItem_Src in $spItems_Src) {
       # Write-Host "Title " $Title_Src 
        $Title_Src = $spItem_Src["Title"]
        #$Title_Src.WebLastScan
        $queryXML = '<View><Query><Where><Eq><FieldRef Name="Title"/><Value Type="Text">'+$Title_Src+'</Value></Eq></Where></Query></View>'
        $spItems_Dest = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Dest -listTitle $ListTitle_Dest -queryXML $queryXML

        If ($spItems_Dest.Count -eq 0) {
           $spNewItem = f_CSOM_Item_Add -spCtx $spCtx_Dest -listTitle $ListTitle_Dest -itemTitle $Title_Src
           Write-host "New" $spNewItem.title 
           $spItems_Dest = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Dest -listTitle $ListTitle_Dest -queryXML $queryXML
        }

        If ($spItem_Src["ID"] -ne $spItems_Dest["WebData"].LookupID) {
           # Write-Host "Title_Src $Title_Src " $spItem_Src["ID"] $spItems_Dest["ID"]
            #$spItem_Src["ID"] 
            #$spItems_Dest["ID"]
            #Write-Host "Mismatch"
            $dataHASH = @{
                "WebData" = $spItem_Src["ID"]
                #"SCData" = $spItem_Src["ID"]
            }
            Write-Host "Title_Src $Title_Src"

            f_CSOM_Item_UpdateByTitle -spCtx $spCtx_Dest -ListTitle $ListTitle_Dest -ItemTitle $Title_Src -spItemFieldsHash $dataHASH
        }
    }
    #>
}

## MAIN ##############################

f_SiteMgmt_dataWebs_Update