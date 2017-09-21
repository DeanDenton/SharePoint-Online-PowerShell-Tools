#http://www.sharepointdiary.com/2017/06/sharepoint-online-how-to-change-list-to-new-experience.html
#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
   
##Variables for Processing
$SiteUrl = "https://crescent.sharepoint.com/sites/sales"
$ListName= "Project Documents"
$UserName="salaudeen@crescent.com"
$Password ="Password goes here"
  
#Setup Credentials to connect
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,(ConvertTo-SecureString $Password -AsPlainText -Force))
 
Try { 
    #Set up the context
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl) 
    $Context.Credentials = $credentials
 
    #Get the document library
    $List = $Context.web.Lists.GetByTitle($ListName)
    #Set list properties
    $List.ListExperienceOptions = "NewExperience" #ClassicExperience or Auto
    $List.Update()
 
    $context.ExecuteQuery()
 
    Write-host "UI experience updated for the list!" -ForegroundColor Green        
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}