#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
   
##Variables for Processing
$SiteUrl = "https://crescent.sharepoint.com/Projects"
$UserName="salaudeen@crescent.com"
$Password ="Password goes here"
  
#Setup Credentials to connect
$Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($UserName,(ConvertTo-SecureString $Password -AsPlainText -Force))
 
Try { 
    #Set up the context
    $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl) 
    $Context.Credentials = $credentials
 
    #Create new document library
    $ListInfo = New-Object Microsoft.SharePoint.Client.ListCreationInformation
    $ListInfo.Title = "Project Documents"
    $ListInfo.TemplateType = 101 #Document Library
    $List = $Context.Web.Lists.Add($ListInfo)
     
    #Set "New Experience" as list property
    $List.ListExperienceOptions = "NewExperience" #Or ClassicExperience
    $List.Update()
    $Context.ExecuteQuery()
 
    Write-host "New Document Library Created!" -ForegroundColor Green
}
catch {
    write-host "Error: $($_.Exception.Message)" -foregroundcolor Red
}