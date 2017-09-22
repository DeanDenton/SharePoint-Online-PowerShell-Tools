#f_CSOM_WF.psm1
#
#Function f_CSOM_WF_ () {}
##################################
Function f_CSOM_WF_GetAllLists ($spCtx, $CSVPath) {
	$Lists=$spCtx.Web.Lists
	$spCtx.Load($spCtx.Web)
	$spCtx.Load($spCtx.Web.Webs)
	$spCtx.Load($Lists)
	$spCtx.ExecuteQuery()
	Foreach ( $ll in $Lists){
		$workflo = $ll.WorkflowAssociations;
		$spCtx.Load($workflo);
		$spCtx.ExecuteQuery();
		Write-host $ll.Title $workflo.Count -ForegroundColor Green 

		foreach ($workfloek in $workflo){
			$workfloek | Add-Member NoteProperty "SiteUrl"($spCtx.Web.Url)
			$workfloek | Add-Member NoteProperty "ListTitle"($ll.Title)
			Write-Output $workfloek
			$workfloek | export-csv $CSVPath -Append
		}
	}
}
##################################





# # # # # # # # # # 
function Get-WorkflowStatus ($spCtx) {
    Add-Type -Path "$CSOM_Path\Microsoft.SharePoint.Client.WorkflowServices.dll"
    $web = $spCtx.Web
    $lists = $web.Lists
	$spCtx.Load($lists);
	$spCtx.ExecuteQuery();

    $workflowServicesManager = New-Object Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager($spCtx, $web);
    $workflowSubscriptionService = $workflowServicesManager.GetWorkflowSubscriptionService();
    $workflowInstanceSevice = $workflowServicesManager.GetWorkflowInstanceService();

    Write-Host ""
    Write-Host "Reviewing Lists" -ForegroundColor Green
    Write-Host ""

    foreach ($list in $lists)        { 
    Write-Host "List" $list.Title
        $workflowSubscriptions = $workflowSubscriptionService.EnumerateSubscriptionsByList($list.Id);
        $spCtx.Load($workflowSubscriptions);                
        $spCtx.ExecuteQuery(); 
        Write-Host "workflowSubscriptions.Count " $workflowSubscriptions.Count               
        foreach($workflowSubscription in $workflowSubscriptions)            {            
            Write-Host "**************************************************************************************"
            Write-Host "List -"$list.Title " Workflow - "$workflowSubscription.Name -ForegroundColor Green
            Write-Host "***************************************************************************************"
            Write-Host ""

            $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
            $camlQuery.ViewXml = "<View> <ViewFields><FieldRef Name='Title' /></ViewFields></View>";
            $listItems = $list.GetItems($camlQuery);
            $spCtx.Load($listItems);
            $spCtx.ExecuteQuery();

            foreach($listItem in $listItems){
                $workflowInstanceCollection = $workflowInstanceSevice.EnumerateInstancesForListItem($list.Id, $listItem.Id);
                $spCtx.Load($workflowInstanceCollection);
                $spCtx.ExecuteQuery();
                foreach ($workflowInstance in $workflowInstanceCollection)
                {
                    Write-Host "List Item Title:"$listItem["Title"] 
                    Write-Host "Workflow Status:"$workflowInstance.Status 
                    Write-Host "Last Updated:"$workflowInstance.LastUpdated
                    Write-Host ""
                }
            }                   
            Write-Host ""
        }
    }  
}