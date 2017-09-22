Function f_CSOM_Ftrs_GetHash () {
    $WebFtrs = @{ 
        "Access" = "d2b9ec23-526b-42c5-87b6-852bd83e0364";
        "Community" = "961d6a9c-4388-4cf2-9733-38ee8c89afd4";
        "ContentOrg" = "7ad5272a-2694-4349-953e-ea5ef290e97c";
        "Following" = "a7a2793e-67cd-4dc1-9fd0-43f61581207a";
        "GettingStarted" = "4aec7207-0d02-4f4f-aa07-b370199cd0c7";
        "MetaDataNav" = "7201d6a4-a5d3-49a1-8c19-19c4bac6e668";
        "MDSFeature" = "87294c72-f260-42f3-a41b-981a2ffce37a";
        "MBrowserRedirect" = "d95c97f3-e528-4da2-ae9f-32b3535fbb59";
        "PremiumWeb" = "0806d127-06e6-447a-980e-2e90b03101b8";
        "PublishingWeb" = "94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb";
        "BaseWeb" = "99fe402e-89a0-45aa-9163-85342e865dc8";
        "SiteFeed" = "15a572c6-e545-4d32-897a-bab6f5846e18";
        "SiteNotebook" = "f151bb39-7c3b-414f-bb36-6bf18872052f";
        "TeamCollab" = "00bfea71-4ea5-48d4-a4ad-7ea5c011abe5";
        "WikiPageHomePage" = "00bfea71-d8fe-4fec-8dad-01c19a6e4053";
        "WorkflowTask" = "57311b7a-9afd-4ff0-866e-9393ad6647b1";
        "SPSBlog" = "d97ded76-7647-4b1e-b868-2af51872e1b3"
        "BlogSiteTemplate" = "faf00902-6bab-4583-bd02-84db191801d8"
        "SitePagesFeatureIdString" = "B6917CB1-93A0-4B97-A84D-7CF49975D4EC"

        #"ProjectFunt" = "?"

    }
    Return $WebFtrs
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Ftrs_ShowAll ($spCtx, $ftrScope) {
    Write-Host ("Show All - Scope: $ftrScope ")
    If ($ftrScope -eq "Web") {
        $spFeatures = $spCtx.Web.Features 
    }
    ElseIf ($ftrScope -eq "Site"){
        $spFeatures = $spCtx.Site.Features 
    }
    $spCtx.Load($spFeatures)
    $spCtx.ExecuteQuery()

    foreach ($spFeature in $spFeatures) {
        Write-Host $spFeature.DefinitionId
    }
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Ftrs_GetAll ($spCtx, $ftrScope) {
    If ($ftrScope -eq "Web") {
        $spFeatures = $spCtx.Web.Features 
    }
    ElseIf ($ftrScope -eq "Site"){
        $spFeatures = $spCtx.Site.Features 
    }
    $spCtx.Load($spFeatures)
    $spCtx.ExecuteQuery()
    Return $spFeatures
}
# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Ftrs_Find ($spCtx, $ftrScope, $ftrGUID) {
    If ($ftrScope -eq "Web") {
        $spFeatures = $spCtx.Web.Features 
    }
    ElseIf ($ftrScope -eq "Site"){
        $spFeatures = $spCtx.Site.Features 
    }
      
    $spCtx.Load($spFeatures)
    $spCtx.ExecuteQuery()

    $ftrFound = $false
    foreach ($spFeature in $spFeatures) {
        If ($ftrGUID -eq $spFeature.DefinitionID) {
            $ftrFound = $true
        }
        #($spFeature.DefinitionID)
    }
    Return $ftrFound
}

# # # # # # # # # # # # # # # # # # # #
Function f_CSOM_Ftrs_Enable($spCtx, $ftrScope, $ftrGUID, $force=$true) {
    $isFtrActive = f_CSOM_Ftrs_Find $spCtx $ftrScope  $ftrGUID
    If ($isFtrActive) {
        $Response="Feature Is already actived: $ftrGUID" 
    }
    Else {
         If ($ftrScope -eq "Web") {
            $spFeatures = $spCtx.Web.Features 
        } 
		ElseIf ($ftrScope -eq "Site"){
            $spFeatures = $spCtx.Site.Features 
        }
        $spCtx.Load($spFeatures)
        $spCtx.ExecuteQuery()
        $Add = $spFeatures.Add($ftrGUID, $force, [Microsoft.SharePoint.Client.FeatureDefinitionScope]::None)
        try {
            $spCtx.ExecuteQuery()
            $Response= "Feature successfully activated: $ftrGUID"
        }
        catch {
            $Response = "An error occurred activating the Feature. Error detail: $($_)"
        }
    }
	Return $Response
}
# # # # # # # # # # # # # # # # # # # 
Function f_CSOM_Ftrs_Disable($spCtx, $ftrScope, $ftrGUID, $force=$true) {
    $isFtrActive = f_CSOM_Ftrs_Find $spCtx $ftrScope $ftrGUID
    If (!$isFtrActive) {
        $Response="Feature Is not active: $ftrGUID"
    }
    Else {
        If ($ftrScope -eq "Web") {
            $spFeatures = $spCtx.Web.Features 
        }ElseIf ($ftrScope -eq "Site"){
            $spFeatures = $spCtx.Site.Features 
        }
        $spCtx.Load($spFeatures)
        $spCtx.ExecuteQuery()
        $spFeatures.Remove($ftrGUID, $force)

        try {
            $spCtx.ExecuteQuery()
            $Response="Feature successfully deactivated: $ftrGUID"
        }
        catch {
            $Response="An error occurred deactivating the Feature. Error detail: $($_)"
        }
    }
	Return $Response
}