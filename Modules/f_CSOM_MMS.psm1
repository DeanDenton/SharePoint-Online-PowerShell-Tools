#Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll"  

#f_CSOM_MMS_

Function f_CSOM_MMS_GetAllTerms($ctx) {
#https://gallery.technet.microsoft.com/scriptcenter/Pull-all-groups-termsets-d489988a
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()

	$session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
	$ctx.Load($session)
	$ctx.ExecuteQuery()

	$termstore = $session.GetDefaultSiteCollectionTermStore();
	$ctx.Load($termstore)
	$ctx.Load($termstore.Groups)
	$ctx.ExecuteQuery()

	Write-Host "Termstore" -ForegroundColor Green
	Write-Output $termstore
	foreach ($gruppo in $termstore.Groups){
		$ctx.Load($gruppo)
		$ctx.Load($gruppo.TermSets)
		$ctx.ExecuteQuery()

		Write-Host "--------------- Group --------------------" -ForegroundColor Yellow
		Write-Output $gruppo 


		foreach($termset in $gruppo.Termsets){
			$ctx.Load($termset)
			$ctx.Load($termset.Terms)
			$ctx.ExecuteQuery()

			Write-Host "--------------- Term Set --------------------" -ForegroundColor Magenta
			Write-Output $termset


			foreach($term in $termset.Terms){
				$ctx.Load($term)
				$ctx.Load($term.Labels)
				$ctx.Load($term.Terms)
				$ctx.Load($term.TermSets)
				$ctx.Load($term.ReusedTerms)
				$ctx.ExecuteQuery()

				Write-Host "--------------- Term --------------------" -ForegroundColor Blue
				Write-Output $term
			}
		}
	}
}


Function f_CSOM_MMS_CreateTerm($ctx,$TermSetGuid,$Term, $TermLanguage=1033) {
#https://gallery.technet.microsoft.com/scriptcenter/Create-a-new-SharePoint-4967719d
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()
	$session = [Microsoft.SharePoint.Client.Taxonomy.TaxonomySession]::GetTaxonomySession($ctx)
	$ctx.Load($session)
	$ctx.ExecuteQuery()

	$termstore = $session.GetDefaultSiteCollectionTermStore();
	$ctx.Load($termstore)
	$ctx.ExecuteQuery()

	$set=$termstore.GetTermSet($TermSetGuid)
	$ctx.Load($set)
	$ctx.Load($set.GetAllTerms())
	$ctx.ExecuteQuery()
	$guid = [guid]::NewGuid()
	$term=$set.CreateTerm($Term, $TermLanguage,$guid)

	$termstore.CommitAll()
	$ctx.ExecuteQuery()
}