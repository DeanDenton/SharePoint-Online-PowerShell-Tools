function Get-SearchProfile () {
    $Scope = "SPSite"
    $Schema = "D:\SearchSchema.XML"

    Add-Type -Path "$CSOM_Path\Microsoft.SharePoint.Client.UserProfiles.dll"
    #Export search configuration
    $Owner = New-Object Microsoft.SharePoint.Client.Search.Administration.SearchObjectOwner($spCtx,$Scope)
    $Search = New-Object Microsoft.SharePoint.Client.Search.Portability.SearchConfigurationPortability($spCtx)
    $SearchConfig = $Search.ExportSearchConfiguration($Owner)
    $Context.ExecuteQuery()
    $SearchConfig.Value > $Schema

}