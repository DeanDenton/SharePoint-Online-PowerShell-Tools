Import-Module ActiveDirectory 

cls

######################################
$CSOM_Path   = "C:\DEV\Microsoft.SharePointOnline.CSOM.16.1.6906.1200\lib\net45\"
$ModulesPath = "C:\DEV\GitHub\SharePoint-Online-PowerShell-Tools\Modules\"
######################################
$ReuseCredentials = $true 
# GET CREDENTIALS  # # # # # # # # # #  
If (($Credentials -eq $Null) -or ($ReuseCredentials -eq $false)) { $Credentials = Get-Credential }

$isSPO = $true
$updateALL = $false
$ScanDateTime = Get-Date
######################################
 # IMPORT #

Add-Type -Path ($CSOM_Path + "Microsoft.SharePoint.Client.dll" )
Add-Type -Path ($CSOM_Path + "Microsoft.SharePoint.Client.Runtime.dll")
######################################
 # MODULES #
import-Module -Name ($ModulesPath + "f_CSOM_Ctx.psm1") -force
import-Module -Name ($ModulesPath + "f_CSOM_List.psm1") -force
import-Module -Name ($ModulesPath + "f_CSOM_Item.psm1") -force
######################################

## FUNCTIONS #########################
Function f_SiteMgmt_dataWebs_ContactADCheck() {


    $SrcURL = "https://maritzllc.sharepoint.com/SUP/WS/"
     $SrcListTitle = "dataWebs" #"Site Management"
    $filterFlag = "SiteMgmtOn"
    $queryXML = '<View><Query><Where><Eq><FieldRef Name="Flag"/><Value Type="Text">'+$filterFlag+'</Value></Eq></Where></Query></View>'

    $spCtx_Src = f_CSOM_Ctx_Get -CtxURL $SrcURL -Credentials $Credentials -isSPO $isSPO
    $spItems_Src = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Src -listTitle $SrcListTitle -queryXML $queryXML

    ForEach ($spItem_Src in $spItems_Src) {
        $itemTitle = $spItem_Src["Title"]
        $AccessRequestEmail = $spItem_Src["AccessRequest"]
        If ($AccessRequestEmail.length -gt 0) {
            If ($AccessRequestEmail -eq "someone@example.com") {
            #Write-Host "$itemTitle - Disabled User: " $AccessRequestEmail -ForegroundColor green
            }
            Else {
            $ADUser = Get-AdUser -Filter{emailaddress -eq $AccessRequestEmail }
            If ($ADUser.Enabled -ne $true) {
            Write-Host "$itemTitle - Disabled Access Request User: " $AccessRequestEmail -ForegroundColor Yellow
                $DataHash = @{
                    "Title" = $itemTitle
                    "OwnerScan" = $AccessRequestEmail
                }
            # f_CSOM_Item_UpdateByTitle -spCtx $spItems_Src -ListTitle $SrcListTitle -ItemTitle $itemTitle -spItemFieldsHash $DataHash   
            }
            }
        }
    }
}
Function f_SiteMgmt_SiteMgmt_ContactADCheck($contactType="Primary") {


    $SrcURL = "https://maritzllc.sharepoint.com/SUP/WS/"
    $SrcListTitle = "Site Management"
    $filterFlag = "SiteDirOn"
    $queryXML = '<View><Query><Where><Eq><FieldRef Name="OnSiteDir"/><Value Type="Text">'+$filterFlag+'</Value></Eq></Where></Query></View>'

    $spCtx_Src = f_CSOM_Ctx_Get -CtxURL $SrcURL -Credentials $Credentials -isSPO $isSPO
    $spItems_Src = f_CSOM_Item_GetAllByQuery -spCtx $spCtx_Src -listTitle $SrcListTitle -queryXML $queryXML

    ForEach ($spItem_Src in $spItems_Src) {
        $itemTitle = $spItem_Src["Title"]
        If ($contactType -eq "Primary") {
            $SiteOwner = $spItem_Src["Site_x0020_Owner"]
        }
        ElseIf ($contactType -eq "Backup") {
            $SiteOwner = $spItem_Src["Site_x0020_Manager"]
        }
        $UserEmail = $SiteOwner.Email 

        If ($UserEmail.length -gt 0) {
            $ADUser = Get-AdUser -Filter{emailaddress -eq $UserEmail }
            If ($ADUser.Enabled -ne $true) {

            Write-Host "$itemTitle - Disabled $contactType Contact: " $UserEmail -ForegroundColor Yellow
                $DataHash = @{
                    "Title" = $itemTitle
                    "OwnerScan" = $UserEmail
                }
            # f_CSOM_Item_UpdateByTitle -spCtx $spItems_Src -ListTitle $SrcListTitle -ItemTitle $itemTitle -spItemFieldsHash $DataHash   

            }
        }
    }
}

f_SiteMgmt_dataWebs_ContactADCheck 
f_SiteMgmt_SiteMgmt_ContactADCheck  -contactType "Backup"
f_SiteMgmt_SiteMgmt_ContactADCheck -contactType "Primary"