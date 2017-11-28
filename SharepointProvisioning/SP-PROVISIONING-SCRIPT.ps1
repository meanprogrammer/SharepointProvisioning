echo "BEGIN SITE PROVISIONING"
#Set-PnPTraceLog -On -Level Debug
Set-PnPTraceLog -On -LogFile traceoutput.txt -Level Debug


#For Dev Purpose only
#$PlainPassword = "plainpassword"
#$SecurePassword = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
#$UserName = "vdudan@adbdev.onmicrosoft.com"
#$credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecurePassword
#For Dev Purpose only

#Get the credential for the execution of script
$credentials = Get-Credential;

#variables
$tenant = "adbdev";
$sourceSite = "/teams/template_collab";
$sourceWebUrl = "https://{0}.sharepoint.com{1}" -f $tenant, $sourceSite;

#TODO: Replace with title and alias of website 
$title = "foo171"
$alias = "foo171"
#TODO: Replace with title and alias of website

#Variables
$sourceWeb = $null
$sourceContext = $null
$web = $null
$context = $null


#Flags: True to generate template
$getAllTemplate = $False;
$getContentTypeTemplate = $False;
$getNavigationTemplate = $False;
$getContentTemplate = $False;
#Flags: True to generate template


try {
    #Connect to the source web (template web)
    Connect-PnPOnline -url $sourceWebUrl -Credentials $credentials
} catch {
    #if there is a problem with the login
    #the script will exit
    echo "Something is wrong with the login."
    Exit
}

#Get the source context and web
$sourceContext = Get-PnPContext
$sourceWeb = $sourceContext.Web
$sourceContext.Load($sourceWeb)
$sourceContext.ExecuteQuery()


echo 'START: GET TEMPLATE'
#if the $getAllTemplate is set to true 
#it will generate the template, else it will skip the generation
#and will assume that the template already exist
if($getAllTemplate -eq $True) {
    Get-PnPProvisioningTemplate -Out "PNP\Complete.pnp" -Force -PersistBrandingFiles -PersistPublishingFiles -IncludeNativePublishingFiles -Handlers All -ExcludeHandlers ComposedLook
} else {
    echo "SKIPPED GET TEMPLATE"
}
echo 'END: GET TEMPLATE'

#if the $getContentTypeTemplate is set to true 
#it will generate the template, else it will skip the generation
#and will assume that the template already exist
echo "START: GETTING CONTENTTYPES"
if($getContentTypeTemplate -eq $True) {
    #Get the field, content types and list template and assign them to variable 
    #for in memory modification
    $tpl = Get-PnPProvisioningTemplate -OutputInstance -Handlers Fields, ContentTypes, Lists -Web $sourceWeb
    $contentTypes = $tpl.ContentTypes
    foreach($ct in $contentTypes){              
           if($ct.Name -eq 'ADB Document' -Or $ct.Name -eq 'ADB Country Document' -Or $ct.Name -eq 'ADB Project Document') {
               #load field reference
               $fields = $ct.FieldRefs
               foreach($ff in $fields) 
               {
                   #set the field to Hidden
                   if($ff.Name -eq 'ADBDocumentTypeValue' -or $ff.Name -eq 'ADBContentGroup') {
                       $ff.Hidden = $True
                   }
               }
           }
    }
    #after the in memory modification, save it to a template file
    Save-PnPProvisioningTemplate -InputInstance $tpl -Force -Out "PNP\collabCT.pnp"

} else {
    echo "SKIPPED GET CONTENT TYPE TEMPLATE"
}
echo "END: GETTING CONTENTTYPES"

#if the $getNavigationTemplate is set to true 
#it will generate the template, else it will skip the generation
#and will assume that the template already exist
echo "START: GETTING NAVIGATION"
if($getNavigationTemplate -eq $True) {
    Get-PnPProvisioningTemplate -Force -Out "PNP\collabNAV.pnp" -Handlers Navigation -Web $sourceWeb
} else {
    echo "SKIPPED GET NAVIGATION TEMPLATE"
}
echo "END: GETTING NAVIGATION"

#if the $getContentTemplate is set to true 
#it will generate the template, else it will skip the generation
#and will assume that the template already exist
echo "START: GETTING HOMEPAGE"
if($getContentTemplate -eq $True) {
    Get-PnPProvisioningTemplate -Force -Out "PNP\collabPP.pnp" -Handlers Pages, PageContents -Web $sourceWeb
} else {
    echo "SKIPPED GET CONTENT TEMPLATE"
}
echo "END: GETTING HOMEPAGE"

#Create the provisioned site
$targetWebUrl = New-PnPSite -Type TeamSite -Title $title -Alias $alias 

#disconnect from the source website
Disconnect-PnPOnline

#connect to the target website
Connect-PnPOnline -url $targetWebUrl -Credentials $credentials;

#get the context, web, lists of the target website
$context = Get-PnPContext
$web = $context.Web
$context.Load($web)
$context.ExecuteQuery()
$context.Load($web.Lists)
$context.ExecuteQuery()

#ensute that site asset library is created
$web.Lists.EnsureSiteAssetsLibrary()
$context.ExecuteQuery()

#ensute that site pages library is created
$web.Lists.EnsureSitePagesLibrary()

$context.ExecuteQuery()

echo "START: TEST IF TARGET IS ACTIVE."
$status = $null
DO

{
    Write-Host "Waiting...   $status"
    Start-Sleep -Seconds 5
    #get the full details of the target site to check if
    #the status flag is now "Active"
    $Site=Get-PnPTenantSite -url $targetWebUrl -Detailed
    $status = $Site.Status

    #continue the loop until the Status is "Active"
} While ($status -ne 'Active')

echo "END: TEST IF TARGET IS ACTIVE."


echo "START: DISABLE NOSCRIPT"
#Set the NoScriptSite flag to false
#this must be set to false for the 
#provisioning of workflow
Set-PnPTenantSite -Url $targetWebUrl -NoScriptSite:$false;
echo "END: DISABLE NOSCRIPT"




#ADD USER AS OWNER TO TARGET
echo "START: ADD USER AS OWNER TO TARGET"
$owners = Get-PnPGroup -AssociatedOwnerGroup -Web $web
Add-PnPUserToGroup -LoginName $credentials.UserName -Identity $owners.Id;
echo "END: ADD USER AS OWNER TO TARGET"




echo "START: APPLY TEMPLATE"
#Apply the template to the target site
Apply-PnPProvisioningTemplate -Path "PNP\Complete.pnp" -Handlers All -ExcludeHandlers ComposedLook -ErrorAction SilentlyContinue 
echo "END: APPLY TEMPLATE"

echo "START: APPLY CONTENTTYPES"
#Apply the content types to the target site
Apply-PnPProvisioningTemplate -Path "PNP\collabCT.pnp" -Handlers Fields, ContentTypes, Lists -Web $web

#Removes the Document content type from the "Final Documents"
Remove-PnPContentTypeFromList -List "Final Documents" -ContentType "Document" -Web $web

#This will remove duplicate fields
Remove-PnPField -List "Documents" -Identity "Update ADB Country Document Type" -Force -Web $web
Remove-PnPField -List "Documents" -Identity "Update ADB Document Type" -Force -Web $web
Remove-PnPField -List "Documents" -Identity "Update ADB Project Document Type" -Force -Web $web
Remove-PnPField -List "Documents" -Identity "Log Activity" -Force -Web $web

echo "END: APPLY CONTENTTYPES"

echo "START: APPLY NAVIGATION"
#Apply the Navigation to the target website
Apply-PnPProvisioningTemplate -Path "PNP\collabNAV.pnp" -ClearNavigation -Handlers Navigation -Web $web
echo "END: APPLY NAVIGATION"

echo "START: APPLY HOMEPAGE"
#Delete existing homepage
Remove-PnPFile -SiteRelativeUrl "SitePages/Home.aspx" -Force -Web $web

#Apply the Pages template
Apply-PnPProvisioningTemplate -Path "PNP\collabPP.pnp" -Handlers Pages, PageContents -Web $web

#adds the News webpart to the homepage
$page = Get-PnPClientSidePage -Identity "Home.aspx" -Web $web
Add-PnPClientSideWebPart -Page $page -DefaultWebPartType NewsFeed -Section 2 -Column 1 -Web $web
echo "END: APPLY HOMEPAGE"

echo "START: ENABLE NOSCRIPT"
#Set the NoScriptSite flag to true
Set-PnPTenantSite -Url $targetWebUrl -NoScriptSite:$true
echo "END: ENABLE NOSCRIPT"

echo "END HOMEPAGE"

#Updates the workflow references
echo "START: UPDATE WORKFLOW REFERENCES"

#Gets the site of the target website
$site = $context.Site
$context.Load($site)
$context.ExecuteQuery()

#gets reference to the "Documents" document library
$documents = $web.Lists.GetByTitle("Documents")

$context.Load($documents)
$context.ExecuteQuery()

#loads all workflow associations
$context.Load($documents.WorkflowAssociations)
$context.ExecuteQuery()

#This is required to use the WorkflowServicesManager class
Add-Type -Path "Microsoft.SharePoint.Client.WorkflowServices.dll"

#Gets the WorkflowServicesManager instance
$servicesManager = New-Object Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager($context, $web)
#Gets the WorkflowSubscriptionService
$subscriptionService = $servicesManager.GetWorkflowSubscriptionService()
#List all the subscription
$subscriptions = $subscriptionService.EnumerateSubscriptionsByList($documents.Id)

$context.Load($subscriptions)
$context.ExecuteQuery()

#Gets a reference to the Workflow history list
$wfh = $web.Lists.GetByTitle("Workflow History");
$wft = $web.Lists.GetByTitle("Update Document Type Workflow Tasks")

#Gets a reference to the Workflow history task
$dwt = $web.Lists.GetByTitle("Workflow Tasks");
$context.Load($wfh)
$context.Load($wft)
$context.Load($dwt)
$context.ExecuteQuery()

#Loop through all the subscription and
#set the HistoryListId and TaskListId
#and publish
foreach ($s in $subscriptions)
{


    if (
        $s.Name -eq "Update ADB Project Document Type" -or
        $s.Name -eq "Update ADB Country Document Type" -or
        $s.Name -eq "Update ADB Document Type"
        )
    {
        $s.SetProperty("HistoryListId", $wfh.Id)
        $s.SetProperty("TaskListId", $wft.Id)
        $s.SetProperty("FormData", "")
        $subscriptionService.PublishSubscriptionForList($s, $documents.Id)
    } 
    else
    {
        $s.SetProperty("HistoryListId", $wfh.Id)
        $s.SetProperty("TaskListId", $dwt.Id)
        $s.SetProperty("FormData", "")
        $subscriptionService.PublishSubscriptionForList($s, $documents.Id)
    }
}
$context.ExecuteQuery()

echo "END: UPDATE WORKFLOW REFERENCES"

$setHome = $FALSE;

#this sets the PageLayoutType of the Home.aspx
#so the banner of the page will be removed
do {
    try
    {

        $web.Lists.EnsureSitePagesLibrary()

        $context.ExecuteQuery()

        echo "get list item"
        $retval = Get-PnPFile -Url "SitePages/Home.aspx" -AsListItem -Web $web #Get-PnPListItem -List "SitePages" #-Query  "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Home</Value></Eq></Where></Query></View>"
        
        echo "set type to home"
        Set-PnPListItem -List "SitePages" -Identity $retval.Id -Values @{"PageLayoutType"="Home"} -Web $web -ErrorAction SilentlyContinue
        echo "set type to home"
        $setHome = $TRUE;
    }   
    catch
    {
        $setHome = $FALSE;
    }

} While($setHome -eq $FALSE)
echo "exited loop"

#disconnect
Disconnect-PnPOnline
#exit script
Exit

