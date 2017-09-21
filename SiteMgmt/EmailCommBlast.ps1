 # Variables
$baseURL = "https://maritzllc.sharepoint.com"
$SrcURL = ($baseURL + "/SUP/WS/")
$SrcListTitle = "Site Management"

$bTesting = $false

$SMTP = "MIFENMAIL99.maritz.com"
$sFrom = "sharepointsupport@maritz.com"
$sBCC = ""

$ReuseCredentials = $true 
$sTimeStamp = Get-Date
######################################
$CSOM_Path = "C:\SYNC\SharePoint\MSH Development - Code\dllCSOM\"
Add-Type -Path ($CSOM_Path + "Microsoft.SharePoint.Client.dll" )
Add-Type -Path ($CSOM_Path + "Microsoft.SharePoint.Client.Runtime.dll")
######################################
# # # # # # # # # # # # # # # # # # #
Function f_CSOM_Ctx_Get($CtxURL, $Credentials=$NULL, $getCredAlways=$false, $isSPO=$true) {
	If (($getCredAlways -eq $true) ){
		$Credentials = Get-Credential
	}
    $spCtx = New-Object Microsoft.SharePoint.Client.ClientContext($CtxURL)  -ErrorAction stop

	$CtxLogin = $Credentials.UserName
	$CtxPwd = $Credentials.Password 

    If ($isSPO) {
        $spCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($CtxLogin, $CtxPwd)   
        $spCtx.Credentials = $spCredentials  
    } Else {
        $spCtx.Credentials = New-Object System.Net.NetworkCredential($CtxLogin, $CtxPwd)
    }
    Return $spCtx    
}
# # # # # # # # # # # # # # # # # # #
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
 # # # # # # # # # # # # # # # # # # #
 # Send Email  
Function f_SendEmail($sFrom, $sTo, $sCC, $sBCC, $sSubject, $sHTMLBody) {
    If ($sCC -ne "") {
        If ($sBCC -ne "") {
            send-mailmessage  -From $sFrom -To $sTo -Cc $sCC -Bcc $sBCC -Subject $sSubject -BodyAsHtml $sHTMLBody -smtp $SMTP 
        } Else {
            send-mailmessage  -From $sFrom -To $sTo -Cc $sCC -Subject $sSubject -BodyAsHtml $sHTMLBody -smtp $SMTP 
        } 
    } Else {
        If ($sBCC -ne "") {
            send-mailmessage  -From $sFrom -To $sTo -Bcc $sBCC -Subject $sSubject -BodyAsHtml $sHTMLBody -smtp $SMTP 
        } Else {
            send-mailmessage  -From $sFrom -To $sTo -Subject $sSubject -BodyAsHtml $sHTMLBody -smtp $SMTP 
        } 
    }
}

# # # # # # # # # # # # # # # # # # #
Function f_SiteMgmt_SiteList ($sSubject, $sBody, $sSendFilter="", $iCommID="") {
    $spItems = f_CSOM_Item_GetAllByQuery -spCtx $spCtx -listTitle $SrcListTitle -queryXML $sSendFilter

    ForEach ($spItem_Site in $spItems) {
        #Get Primary Contact
        $UserValue = $spItem_Site["Site_x0020_Owner"]
        $sTo = $UserValue.Email
        $PrimaryContact = $UserValue.LookupValue

        #Get Backup Contact
        $UserValue = $spItem_Site["Site_x0020_Manager"]
        If ($UserValue.LookupId -gt 0) {
            $sCC = $UserValue.Email
            $BackupContact = $UserValue.LookupValue
        }
        Else {
            $sCC = ""
            $BackupContact = " - Backup Contact Needed - "
        }

        #Get Web Data
        $webListTitle = "dataWebs"
        $queryXML = ('<View><Query><Where><Eq><FieldRef Name="ID"/><Value Type="Int">'+$spItem_Site["WebData"].LookupID+'</Value></Eq></Where></Query></View>')
        $spWebItem = f_CSOM_Item_GetAllByQuery -spCtx $spCtx -listTitle $webListTitle -queryXML $queryXML

        $subject = $sSubject + " (" + $spItem_Site["Title"] + ")"

        #InsertSiteDataHere 
        $SiteDataInsert = ("<b>Site Name:</b> <a href='"+$spWebItem["WebUrl"]+"'>"+ $spWebItem["WebTitle"] +"</a><br/><b>Site Path:</b> " + $baseURL + $spItem_Site["Title"] + "<br/><b>Primary Contact:</b> " + $PrimaryContact + "<br/><b>Backup Contact:</b> " + $BackupContact )

        $HTMLBody = $sBody.Replace("InsertSiteDataHere", $SiteDataInsert)

        If ($bTesting -eq $false) {
            Write-Output ("Send Email To: " + $spItem_Site["Title"] + "|" + $PrimaryContact  + "|" + $sTo  + "|" + $iCommID )

            $spItem_Site["LatestComm"] = $iCommID 
            $spItem_Site.Update() 
            $spCtx.ExecuteQuery() 
        } Else {
            $sTo = "Dean.Denton@Maritz.com"
            $sCC = ""
            $sBCC = ""
            $spItem_Site["Title"] + "|" + $PrimaryContact  + "|" + $sTo  + "|" + $iCommID
        }

        f_SendEmail -sFrom $sFrom -sTo $sTo -sCC $sCC -sBCC $sBCC -sSubject $subject -sHTMLBody ($HTMLBody) 
    }
}

# # # # # # # # # # # # # # # # # # #
Function f_SiteMgmt_CommList () {
    $CommListTitle = "Site Communications"
    $queryXMLNewOnly = '<View><Query><Where><And><Eq><FieldRef Name="Status"/><Value Type="Text">Ready</Value></Eq><IsNull><FieldRef Name="DateSent"/></IsNull></And></Where></Query></View>'
    $spItems = f_CSOM_Item_GetAllByQuery -spCtx $spCtx -listTitle $CommListTitle -queryXML $queryXMLNewOnly

    ForEach ($spItem_Comm in $spItems) {
        $SendFilter = $spItem_Comm["SendFilter"]
        If ($SendFilter -ne "") {
            f_SiteMgmt_SiteList -sSubject $spItem_Comm["Subject"] -sBody $spItem_Comm["Body"] -sSendFilter $SendFilter -iCommID $spItem_Comm["ID"]
        }
        If ($bTesting -eq $false) {  
            $sTimeStamp       
            $spItem_Comm["DateSent"] = $sTimeStamp 
            $spItem_Comm["Status"] = "Complete"
            $spItem_Comm.Update() 
            $spCtx.ExecuteQuery() 
        }
    }
}

#MAIN ##########################

Start-Transcript -Path "C:\SYNC\DEV\Output\Transcripts\EmailBlast.txt" -Force

# GET CREDENTIALS  # # # # # # # # # #  
If (($Credentials -eq $Null) -or ($ReuseCredentials -eq $false)) { $Credentials = Get-Credential }

$spCtx = f_CSOM_Ctx_Get -CtxURL $SrcURL -Credentials $Credentials 

f_SiteMgmt_CommList

Stop-Transcript 

#Future: Attach Transcript


        #$queryXML = ''

        #$queryXML = '<View><Query><Where><And><Eq><FieldRef Name="OnSiteDir"/><Value Type="Text">SiteDirOn</Value></Eq><IsNull><FieldRef Name="OwnerConfirmed"/></IsNull></And></Where><OrderBy><FieldRef Name="Modified" Ascending="FALSE"/></OrderBy></Query></View>'
        
        #$queryXML = '<View><Query><Where><Eq><FieldRef Name="OnSiteDir"/><Value Type="Text">SiteDirOn</Value></Eq></Where><OrderBy><FieldRef Name="Title" Ascending="TRUE"/></OrderBy></Query></View>'

        