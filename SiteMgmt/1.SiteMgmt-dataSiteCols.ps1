#d2.Note: Add functionality to send notification of site w/o uso_Exchange_SharePointOnline
#d2.note: Get diff between Site, Group, Video
######################################

 # CONTROL #
$TenantAdminURL = "https://maritzllc-Admin.sharepoint.com/"
$ControlURL = "https://maritzllc.sharepoint.com/SUP/ws/"
$ControlListTitle = "dataSiteCols"
$SCLastScan = Get-Date

$ReuseCredentials = $true

If (($Credentials -eq $Null) -or ($ReuseCredentials -eq $false)) { 
    # GET TENANT #
    Connect-SPOService -Url $TenantAdminURL #-Credential $Credentials
    $Credentials = Get-Credential 
}

######################################
$CSOM_Path   = "C:\DEV\Microsoft.SharePointOnline.CSOM.16.1.6906.1200\lib\net45\"
$ModulesPath = "C:\DEV\GitHub\SharePoint-Online-PowerShell-Tools\Modules\"
######################################
 # IMPORT #
Add-Type -Path "$CSOM_Path\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "$CSOM_Path\Microsoft.SharePoint.Client.Runtime.dll" 
Add-Type -Path "$CSOM_Path\Microsoft.Online.SharePoint.Client.Tenant.dll"
#Add-Type -Path "$CSOM_Path\Microsoft.Online.SharePoint.Client.Tenant.dll.15.0.4615.1001\Microsoft.Online.SharePoint.Client.Tenant.dll"
######################################
 # MODULES #
import-Module -Name ($ModulesPath + "f_CSOM_Ctx.psm1") -force
import-Module -Name ($ModulesPath + "f_CSOM_List.psm1") -force
import-Module -Name ($ModulesPath + "f_CSOM_Item.psm1") -force
#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-#-# 
#Use Gridview as selection tool
# $sitecollections = Get-SPOSite -Limit ALL -Detailed | Out-GridView -Title "Select site collections from which to collect data." -PassThru;

######################################
 # MAIN #

Function f_SiteMgmt_dataSiteCols() {
    $spCtx_Control = f_CSOM_Ctx_Get -CtxURL $ControlURL -Credentials $Credentials -getCredAlways $false -isSPO $true

    # GET CONTROL LIST ITEMS # # # # # # # # # # # #  
    #$queryXML = '<View><Query><Where><Eq><FieldRef Name="Site_x0020_Type"/><Value Type="Text">Site</Value></Eq></Where></Query></View>'
    $spControlItems  = f_CSOM_Item_GetAllByListTitle -spCtx $spCtx_Control -listTitle $ControlListTitle #-queryXML  $queryXML 

    $ControlSiteArray = @();
    ForEach($spItem in $spControlItems ) {
        #$SiteURLS = New-Object System.Object;
        #$SiteURLS | Add-Member -MemberType NoteProperty -Name "Url" -Value $spItem["Title"] 
        $ControlSiteArray +=   $spItem["Title"]
    }
    #$ControlSiteArray |Out-Gridview

    #Out-GridView Passthrough allows selection
    $sitecollections = get-sposite  -Limit All  #| Out-GridView -Title "Select site collections from which to collect data." -PassThru;
    #$sitecollections | out-gridview

    $itemsArray = @();
    foreach($item in $sitecollections){
        $itemTitle = $item.Url.Replace("https://maritzllc.sharepoint.com","")

#If ($itemTitle -eq "/sites/BITS" ) {
        If ($ControlSiteArray -notcontains $itemTitle ) {
            #NEW
            write-host "Creating New Item " $itemTitle "=" $item.Url -ForegroundColor green
            $spItem = f_CSOM_Item_Add -spCtx $spCtx_Control -listTitle $ControlListTitle  -itemTitle $itemTitle 
        }
        #$item | Out-GridView

        Try {
            $siteDetails = Microsoft.Online.SharePoint.Powershell\get-sposite -Identity $item.Url -Detailed
            $webCount = $siteDetails.WebsCount
        }
        Catch {
            Write-Host ("Bad Site Format " + $item.Url ) -foregroundcolor cyan
            $webCount = -1
        }

        Try {
            $spCtxSite = f_CSOM_Ctx_Get -CtxURL $item.Url -Credentials $Credentials -getCredAlways $false -isSPO $true
            $spCtxSite.Load($spCtxSite.Web.RegionalSettings.TimeZone)
            $spCtxSite.ExecuteQuery()
            $TimeZone = $spCtxSite.Web.RegionalSettings.TimeZone.Description.tostring()
        }
        Catch {
            $TimeZone = "Unknown"

        }
        
        #UPDATE
        $itemArray = @{
            "SiteColURL" = $item.Url;
            "SiteColTitle" = $item.Title;
            "SiteColOwner" = $item.Owner;
            "WebsCount" = $webCount
            "UsageMB" = $item.StorageUsageCurrent;
            "QuotaMB" = $item.StorageQuota;
            "Sharing" = $item.SharingCapability
            "LastModified" = $item.LastContentModifiedDate
            "LastScan" = $scLastScan
            "TimeZone" = $TimeZone
            #Site.Features
          }
        $timeStamp = Get-Date -format HH:mm:ss
        Write-Host ( $timeStamp.ToString() + "| Updating " + $item.Url) -ForegroundColor Yellow
        f_CSOM_Item_UpdateByTitle -spCtx $spCtx_Control -ListTitle $ControlListTitle -ItemTitle $itemTitle -spItemFieldsHash $itemArray
#}
    }
    # Create Array of New - Item Created
    # Create Array of Removed - SCLastScan < Today
}

f_SiteMgmt_dataSiteCols

Write-Host "Succesfully updated site collection information." -ForegroundColor Green