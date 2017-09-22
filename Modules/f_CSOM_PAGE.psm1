# # # # # # # # # # # # # # # # # # #
Function f_CSOM_PAGE_GetWelcomePage($spCtx) {
    $rootFolder = $spCtx.Web.RootFolder
    $spCtx.Load($rootFolder)
    $spCtx.ExecuteQuery();
    Return $rootFolder.WelcomePage
}
# # # # # # # # # # # # # # # # # # #
function CheckPageExists(){
#https://gist.github.com/asadrefai/20ebe22656a04996453b
    param(
        [Parameter(Mandatory=$true)][string]$siteurl,
        [Parameter(Mandatory=$false)][System.Net.NetworkCredential]$credentials,
        [Parameter(Mandatory=$false)][string]$PageName
    )
    begin{
        try
        {
            $context = New-Object Microsoft.SharePoint.Client.ClientContext($siteurl)
	        $context.Credentials = $credentials

	        
	        $Pages = $context.Web.Lists.GetByTitle('Pages')
	        $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
	        $camlQuery.ViewXml = '<View><Query><Where><Eq><FieldRef Name="FileLeafRef" /><Value Type="Text">'+ $PageName +'</Value></Eq></Where></Query></View>'
	        $Page = $Pages.GetItems($camlQuery)
	        $context.Load($Page)
	        $context.ExecuteQuery()
        }
        catch
        {
            Write-Host "Error while getting context. Error -->> "  + $_.Exception.Message -ForegroundColor Red
        }
    }
    process{
        try
        {
            $file = $Page.File
            if($file)
            {
                #Execute if Page exists
                Write-Host "Page exists" -ForegroundColor Green
                return $true
            }
            else
            {
                #Execute if Page does not exists
                Write-Host "Page does not exists" -ForegroundColor Green
                return $false
            }

        }
        catch
        {
            Write-Host ("Error while checking for Page. Error -->> " + $_.Exception.Message) -ForegroundColor Red
        }
    }
    end{
        $context.Dispose()
    }
}
# # # # # # # # # # # # # # # # # # #
function ChangePageLayout(){
    param(
        [Parameter(Mandatory=$true)][string]$siteurl,
        [Parameter(Mandatory=$false)][System.Net.NetworkCredential]$credentials,
        [Parameter(Mandatory=$false)][string]$PageName,
        [Parameter(Mandatory=$false)][string]$PageLayoutName,
        [Parameter(Mandatory=$false)][string]$PageLayoutDisplayName,
        [Parameter(Mandatory=$false)][string]$Title,
        [Parameter(Mandatory=$false)][bool]$isCustomPageLayout
    )
	#https://gist.github.com/asadrefai/d59cb8d58f1602df2888 
    try
    {

        $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteurl)
	    $ctx.Credentials = $credentials

        if($isCustomPageLayout -eq $false)
        {
            $PageLayoutName = "/_catalogs/masterpage/" + $PageLayoutName + "," + $PageLayoutDisplayName
        }
        else
        {
            #Here I have assumed that if its custom page layout, then it's placed inside some folder which is child to masterpage
            #If that's not the case with you then you can use below line of code
            #$PageLayoutName = "/_catalogs/masterpage/" + $PageLayoutName + ", " + $PageLayoutDisplayName
            #
            $PageLayoutName = "/_catalogs/masterpage/Custom Page Layouts/" + $PageLayoutName + ", " + $PageLayoutDisplayName
        }
	    
	    $Pages = $ctx.Web.Lists.GetByTitle('Pages')
	    $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
	    $camlQuery.ViewXml = '<View><Query><Where><Eq><FieldRef Name="FileLeafRef" /><Value Type="Text">'+ $PageName +'</Value></Eq></Where></Query></View>'
	    $Page = $Pages.GetItems($camlQuery)
	    $ctx.Load($Page)
	    $ctx.ExecuteQuery()
	    
	    $file = $Page.File

	    $ctx.Load($file)
	    $ctx.ExecuteQuery()

        if ($file.CheckOutType  -ne [Microsoft.SharePoint.Client.CheckOutType]::None) {
            $file.UndoCheckOut()
	        $ctx.Load($file)
	        $ctx.ExecuteQuery() 
        }

	    $file.CheckOut()
	    $ctx.Load($file)
	    $ctx.ExecuteQuery()
	    
	    $Page.Set_Item("PublishingPageLayout", $PageLayoutName)
        $Page.Set_Item("Title", $Title)
        $Page.Update()
	    $Page.File.CheckIn("", [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
	    $Page.File.Publish("")

        #check for approval
        $ctx.Load($Pages)
        $ctx.ExecuteQuery()

        if ($Pages.EnableModeration -eq $true) {
            $Page.File.Approve("")
        }

	    $ctx.Load($Page)
	    $ctx.ExecuteQuery()
	    Write-Host "Update Page Layout Complete"
	    Write-Host ""

    }
    catch
    {
        Write-Host ("Error while changing page layout. Error -->> " + $_.Exception.Message) -ForegroundColor Red
    }
}
# # # # # # # # # # # # # # # # # # #