
######################################
Function f_CSOM_Perm_GetUser($spCtx, $User) {
  	$spUser = $spCtx.Web.EnsureUser($User) 
	$spCtx.Load($spUser) 
    $spCtx.ExecuteQuery()
    Return $spUser 
 }######################################
Function f_CSOM_Perm_GetUserLoginName ($spCtx, $User) {
  	$spUser = $spCtx.Web.EnsureUser($User) 
	$spCtx.Load($spUser) 
    $spCtx.ExecuteQuery()
	#Write-Host "spUser.LoginName " $spUser.LoginName
    Return $spUser.LoginName
 }
 ######################################
Function f_CSOM_Perm_Group_GetByName($spCtx, $GroupName) {
#Write-Host  "CSOM_Perm_Group_GetByName($spCtx, $GroupName)"
	$spGroups = $spCtx.Web.SiteGroups
	$spCtx.Load($spGroups)
	$spCtx.ExecuteQuery()
	$spGroup = $spGroups.GetByName($GroupName)
	$spCtx.Load($spGroup)
	Try {$spCtx.ExecuteQuery()} 
	Catch { Write-host ($GroupName + " " + $GroupName + " Does Not Exist") -foregroundcolor red }
	Return $spGroup
}
######################################
Function f_CSOM_Perm_Group_Create($spCtx, $GroupName, $GroupDescription="") {
	$spGroup = f_CSOM_Perm_Group_GetByName -spCtx $spCtx -GroupName $GroupName 
    If ($spGroup.Id -gt 0 ) {
    }
    Else {
	    $spGroupInfo=New-Object Microsoft.SharePoint.Client.GroupCreationInformation 
	    $spGroupInfo.Title=$GroupName 
	    $spGroupInfo.Description=$GroupDescription 
	    $spGroup=$spCtx.Web.SiteGroups.Add($spGroupInfo) 
	    $spCtx.ExecuteQuery()
        $spGroup = f_CSOM_Perm_Group_GetByName -spCtx $spCtx -GroupName $GroupName
    }
    	
	Return $spGroup 
}

######################################
Function f_CSOM_Perm_CheckforNullUser($spCtx, $User) {
    If ($User -ne $Null) {
        $contact = f_CSOM_Perm_GetUser -spCtx $spCtx -User ($User.Email)
    } Else { 
        $contact = $Null
    }
    #$contact | out-gridview
    #Write-host "Contact Email " $contact.Email
    Return $contact
} 

######################################
Function f_CSOM_Perm_Group_Delete($spCtx, $GroupName){
	$spGroup = f_CSOM_Perm_Group_GetByName -spCtx $spCtx -GroupName $GroupName 
	Try { 
		$spCtx.Web.SiteGroups.Remove($spGroup)
		$spCtx.ExecuteQuery()
		("Group Delete Succeeded: " + $spGroup.Title + " Deleted") 
	}
	Catch { 
		("Group Delete Failed: " + $GroupName + " Does Not Exist")	
	}
}
######################################
Function f_CSOM_Perm_Group_SetOwner($spCtx, $GroupName, $OwnerName) { #Does Not Work

	$spGroup = f_CSOM_Perm_Group_GetByName -spCtx $spCtx -GroupName $GroupName 
	$spOwner = f_CSOM_Perm_Group_GetByName -spCtx $spCtx -GroupName $OwnerName 
	$spGroup.Owner = $spOwner
	$spGroup.Update()
	$spCtx.ExecuteQuery()
}
######################################
Function f_CSOM_Perm_Group_AddMember($spCtx, $GroupName, $MemberToAdd) {
	$spGroup = f_CSOM_Perm_Group_GetByName -spCtx $spCtx -GroupName $GroupName 
	$Member = $spCtx.Web.EnsureUser($MemberToAdd) 
	$spCtx.Load($Member) 
	$addMember=$spGroup.Users.AddUser($Member) 
	$spCtx.Load($addMember) 
	Try{
		$spCtx.ExecuteQuery()
		$Response = ("Member Added:" + $MemberToAdd)
	}
	Catch {		
		$Response = ("Member Add Error:" + $MemberToAdd )	
	}
	Return $Response
}
######################################

Function f_CSOM_Perm_WebInheritance ($spCtx, $bInheritanceOn, $bResetFirst=$false) {
    $spWeb = $spCtx.Web
    $spCtx.Load($spWeb)
    $spCtx.ExecuteQuery()
	Try {	
		If ($bInheritanceOn) {
			$spWeb.ResetRoleInheritance()
		}
		ElseIf (-not $bInheritanceOn ) {
			If ($bResetFirst) {
				$spWeb.ResetRoleInheritance()
				$spWeb.Update()
				$spCtx.ExecuteQuery()
			}
			$spWeb.breakRoleInheritance($false,$false)
		}
		$spWeb.Update()

		$spCtx.ExecuteQuery()
		$Response = ("Inheritance On set to " + $bInheritanceOn.ToString() )
	}
	Catch {
		$Response = ("Inheritance Unable to Be set to " + $bInheritanceOn.ToString() )
	}
    Return $Response
}
######################################
Function f_CSOM_Perm_WebPerms ($spCtx ) {
    $spRoleAssignments = $spCtx.Web.RoleAssignments
    $spCtx.Load($spRoleAssignments)
    $spCtx.ExecuteQuery()

    ForEach($spRoleAssignment in $spRoleAssignments){
        $spMember = $spRoleAssignment.Member
        $spCtx.Load($spMember)
        $spCtx.ExecuteQuery()
        Write-Host "Role: " $spMember.Title " - " $spMember.PrincipalType  -ForegroundColor Green   

        $spRoleDefBindings = $spRoleAssignment.RoleDefinitionBindings
        $spCtx.Load($spRoleDefBindings)
        $spCtx.ExecuteQuery()

        ForEach ($spRoleDefBinding in $spRoleDefBindings)  {
            Write-Host "RoleDefBinding Id:" $spRoleDefBinding.Id  -ForegroundColor DarkYellow
            Write-Host  "RoleDefBinding Name:" $spRoleDefBinding.Name  -ForegroundColor DarkYellow
        }
        Write-Host
    }
}
######################################
Function f_CSOM_Perm_Group_WebAssoc($spCtx, $GroupName, $AssocGroup) {
	$spGroup = f_CSOM_Perm_Group_GetByName -spCtx $spCtx -GroupName $GroupName 
	switch ($AssocGroup) {
		"Visitor" 	{$spCtx.Web.AssociatedVisitorGroup 	= $spGroup}
		"Member"	{$spCtx.Web.AssociatedMemberGroup 	= $spGroup}
		"Owner"		{$spCtx.Web.AssociatedOwnerGroup 	= $spGroup}
	}
	$spCtx.Web.Update()
	Try {
		$spCtx.ExecuteQuery()
		$Response = "Group Associated with $AssocGroup"
	}
	Catch {
		$Response = "Group Association Error"
	}
	Return $Response
}
######################################
Function f_CSOM_Perm_Group_GetWebAssoc($spCtx, $AssocGroup) {
	switch ($AssocGroup) {
		"Visitor" 	{$spGroup = $spCtx.Web.AssociatedVisitorGroup }
		"Member"	{$spGroup = $spCtx.Web.AssociatedMemberGroup}
		"Owner"		{$spGroup = $spCtx.Web.AssociatedOwnerGroup }
	}
	Try {
		$spCtx.ExecuteQuery()
		$Response = $spGroup
	}
	Catch {
		$Response = $Null
	}
	Return $Response
}
######################################
Function f_CSOM_Perm_AssignPerm($spCtx, $MemberName, $MemberType, $PermName) {
	If ($MemberType -eq "Group") {
		$spMember = f_CSOM_Perm_Group_GetByName -spCtx $spCtx -GroupName $MemberName 
	} Else {
		$spMember = $spCtx.Web.EnsureUser($MemberName) 
	}
	
	#Write-Host "spMember $spMember"
	
	$PermissionLevel = $spCtx.Web.RoleDefinitions.GetByName($PermName)
	$RoleDefBind = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($spCtx)
	$RoleDefBind.Add($PermissionLevel)
	$Assignments = $spCtx.Web.RoleAssignments
	$RoleAssign = $Assignments.Add($spMember,$RoleDefBind)
	
	$spCtx.Load($spMember)
	$spCtx.ExecuteQuery()
	Try {	
		$spCtx.ExecuteQuery()	
		$Response = "Permission Assigned $MemberName => $PermName"
	}
	Catch { 
		$Response = ("Unable to Assign " + $MemberType + " " + $MemberName + " Permissions " + $PermNameGroup )
	}
	Return $Response
}
######################################
Function f_CSOM_Perm_SiteGroups_GetAll($spCtx) {
    try {
        $spGroups=$spCtx.Web.SiteGroups 
        $spCtx.Load($spGroups) 
        $spCtx.ExecuteQuery() 
        
        foreach($spGroup in $spGroups){ 
            $spCtx.Load($spGroup) 
            $spCtx.ExecuteQuery() 
            Write-Host "spGroup: " $spGroup.Title 
        } 
     } 
    catch [System.Exception]     { 
        write-host -f red $_.Exception.ToString()    
    }
}
######################################
Function f_CSOM_Perm_WebAssignedMembers ($spCtx) {
    #$spWeb = $spCtx.Web
    #$spCtx.Load($spWeb)
    #$spCtx.ExecuteQuery()

    $spRoleAssignments = $spCtx.Web.RoleAssignments
    $spCtx.Load($spRoleAssignments)
    $spCtx.ExecuteQuery()
    ForEach ($spRoleAssignment in $spRoleAssignments) {
        $spRoleMember = $spRoleAssignment.Member
        $spCtx.Load($spRoleMember)
        $spCtx.ExecuteQuery()
        Write-Host ("RoleMember =  " + $spRoleMember.Title + " | " + $spRoleMember.PrincipalType ) -ForegroundColor Yellow
    }
}
######################################
Function f_CSOM_Perm_GroupRemove ($ctx, $groupName) {
	$ctx.Load($ctx.Web)
	$ctx.ExecuteQuery()
	$Groups = $ctx.Web.SiteGroups
	$ctx.Load($Groups)
	$ctx.ExecuteQuery()
	#foreach($group in $Groups){ 		Write-host $group.LoginName	}
	$Groups.RemoveByLoginName($groupName)
	$ctx.Web.Update()
	$ctx.ExecuteQuery()
}

<#
Function f_CSOM_Perm_CreateLevel() {}

Function f_CSOM_Perm_ClearAll () {}

######################################
Function f_CSOM_Perm_Inheritance (){
	#$spWeb.BreakRoleInheritance($true, $false)
	$spList.BreakRoleInheritance($True, $False)
}
######################################

######################################
Function f_CSOM_Perm_AssignPermList() {}
Function f_CSOM_Perm_AssignPermItem() {}
######################################
Function f_CSOM_Perm_GetAll(spCtx) {
	#f_CSOM_Perm_Groups_GetAll
}


Function f_CSOM_Perm_List_Assign($spCtx, $ListName, $MemberName, $MemberType, $PermLevel){
 #http://stackoverflow.com/questions/31112024/adding-read-only-permissions-to-a-list-for-a-sharepoint-group-using-powershell-c
	#function get list by id?
	$spLists = $spCtx.Web .Lists
	$spCtx.Load($spLists)
	$spCtx.ExecuteQuery()
	foreach($spList in $spLists) {
		if($spList.Title -eq $ListName){
			$listId = $spList.Id
		}
	}
	$spList = $spLists.GetById($listId)
	$spCtx.Load($spList);
	$spCtx.ExecuteQuery();
	if ($spList -ne $null) {
		$spGroups = $spCtx.Web.SiteGroups
		$spCtx.Load($spGroups)
		$spCtx.ExecuteQuery()
		foreach ($SiteGroup in $groups) {                    
			if ($SiteGroup.Title -match "Students")
			{
				write-host "Group:" $SiteGroup.Title -foregroundcolor Green
				$GroupName = $SiteGroup.Title

				$builtInRole = $ctx.Web.RoleDefinitions.GetByName($PermissionLevel)

				$roleAssignment = new-object Microsoft.SharePoint.Client.RoleAssignment($SiteGroup)
				$roleAssignment.Add($builtInRole)

				$list.BreakRoleInheritance($True, $False)
				$list.RoleAssignments.Add($roleAssignment)
				$list.Update();
				Write-Host "Successfully added <$GroupName> to the <$ListName> list in <$site>. " -foregroundcolor Green
			}                
			else
			{
					Write-Host "No Students groups exist." -foregroundcolor Red
			}
		}
	}
	

}




#https://social.technet.microsoft.com/Forums/office/en-US/d060dcf6-c827-4343-9f19-536f4f2e7fcb/csom-powershell-code-to-remove-permission-level-from-a-sharepoint-group?forum=sharepointgeneral
#Remove the role
           $roleAssign.RoleDefinitionBindings.Remove($roleDefinition);
           $ctx.Load($roleAssign)
$roleAssign.Update()
$list.Update();

#http://stackoverflow.com/questions/23432665/sharepoint-online-csom-associated-site-groups-associatedmembergroup-associatedow

#Update Group
$currentGroup = $Web.SiteGroups.GetByName($spoGName)
$currentGroup.AllowMembersEditMembership = $false 
$currentGroup.OnlyAllowMembersViewMembership = $true 


            # get role definition
            $roleDefs = $web.RoleDefinitions
            $context.Load($roleDefs)
            $context.ExecuteQuery()
			
            $roleDef = $roleDefs | where {$_.RoleTypeKind -eq $Roletype}
			
# Assign permissions
            $collRdb = new-object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($context)
            $collRdb.Add($roleDef)
            $collRoleAssign = $list.RoleAssignments
            $rollAssign = $collRoleAssign.Add($group, $collRdb)
            $context.ExecuteQuery()
	
			
			
			
			$Web = $Context.Web
$Context.Load($web)
$permissionlevel = "ManageLists, CancelCheckOut, AddListItems, EditListItems, DeleteListItems, ViewListItems, ApproveItems, OpenItems, ViewVersions, DeleteVersions, CreateAlerts, ViewFormPages, ManagePermissions, BrowseDirectories, ViewPages, EnumeratePermissions, BrowseUserInfo, UseRemoteAPIs, Open"
$RoleDefinitionCol = $web.RoleDefinitions
$Context.Load($roleDefinitionCol)
$Context.ExecuteQuery()
$permExists = $false
$spRoleDef = New-Object Microsoft.SharePoint.Client.RoleDefinitionCreationInformation
$spBasePerm = New-Object Microsoft.SharePoint.Client.BasePermissions
$permissions = $permissionlevel.split(",");
foreach($perm in $permissions){$spBasePerm.Set($perm)}

$spRoleDef.Name = $permName
$spRoleDef.Description = $permDescription
$spRoleDef.BasePermissions = $spBasePerm    
$roleDefinition = $web.RoleDefinitions.Add($spRoleDef)
$Context.ExecuteQuery()





#>