[CmdletBinding()]
Param(
   [Parameter(Mandatory=$True)]
   [string]$tenant,

  [Parameter(Mandatory=$True)]
   [string]$alias,
	
   [Parameter(Mandatory=$True)]
   [string]$title,
	
   [Parameter(Mandatory=$True)]
   [bool]$generateTemplate,

   [Parameter(Mandatory=$False)]
   [string]$sourceSite = "/teams/template_pnp"
)

#Functions
#Connect to a SharePoint Online Site
Function Connect-SPOSite
{
    Param([string]$url)
    try {
        #Connect to the source web (template web)
        echo "CONNECTING TO $url"
        Connect-PnPOnline -url $url -Credentials $credentials
    } catch {
        #if there is a problem with the login
        #the script will exit
        $ErrorMessage = $_.Exception.Message
        $FailedItem = $_.Exception.ItemName
        echo "CONNECTION ERROR: $ErrorMessage $FailedItem"
        Exit
    }

}
 


echo "BEGIN SITE PROVISIONING"

#Set-PnPTraceLog -On -Level Debug
Set-PnPTraceLog -On -LogFile traceoutput.txt -Level Error

#Set variables
$sourceSite = "/teams/template_pnp";
$sourceWebUrl = "https://{0}.sharepoint.com{1}" -f $tenant, $sourceSite;
$sourceWeb = $null
$sourceContext = $null
$web = $null
$context = $null

#Flags: True to generate template
$getAllTemplate = $generateTemplate;
$getNavigationTemplate = $generateTemplate;
$getContentTemplate = $generateTemplate;

#Get the credential for the execution of script
$credentials = Get-Credential;


#Check if the target site already exists. Terminate the script if yes, else continue
echo "START: CHECK IF SITE EXISTS"
$targetWebUrl = "https://{0}.sharepoint.com/teams/{1}" -f $tenant, $alias
Try
{
    $existingSite = Get-PnPTenantSite -url $targetWebUrl -Detailed -ErrorAction Stop
    if($existingSite -ne $null)
    {
        echo "Target site already exist. Terminating provisioning script."
        Exit
    }
}
catch
{
    #swallowing exception
    #echo $_.Exception.Message
}


#Call Function Connect-SPOSite to connect to Template Site
Connect-SPOSite -url  $sourceWebUrl


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
    Get-PnPProvisioningTemplate -Out "PNP\Complete.xml" -Force -PersistBrandingFiles -PersistPublishingFiles -IncludeNativePublishingFiles -Web $sourceWeb

} else {
    echo "SKIPPED GET TEMPLATE"
}

echo 'END: GET TEMPLATE'

#if the $getNavigationTemplate is set to true 
#it will generate the template, else it will skip the generation
#and will assume that the template already exist
echo "START: GETTING NAVIGATION"
if($getNavigationTemplate -eq $True) {
    Get-PnPProvisioningTemplate -Force -Out "PNP\collabNAV.xml" -Handlers Navigation -Web $sourceWeb
} else {
    echo "SKIPPED GET NAVIGATION TEMPLATE"
}
echo "END: GETTING NAVIGATION"

#if the $getContentTemplate is set to true 
#it will generate the template, else it will skip the generation
#and will assume that the template already exist
echo "START: GETTING HOMEPAGE"
if($getContentTemplate -eq $True) {
    Get-PnPProvisioningTemplate -Force -Out "PNP\collabPP.xml" -Handlers Pages, PageContents -Web $sourceWeb
} else {
    echo "SKIPPED GET CONTENT TEMPLATE"
}
echo "END: GETTING HOMEPAGE"


#Create the provisioned site
$targetWebUrl = New-PnPSite -Type TeamSite -Title $title -Alias $alias 



Start-Sleep -Seconds 60

#disconnect from the source website
Disconnect-PnPOnline

#connect to the target website
Connect-PnPOnline -Url $targetWebUrl -Credentials $credentials

Start-Sleep -Seconds 60

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
Apply-PnPProvisioningTemplate -Path "PNP\Complete.xml" -Web $web
echo "END: APPLY TEMPLATE"

echo "START: APPLY CONTENTTYPES"
#Removes the Document content type from the "Final Documents"
Remove-PnPContentTypeFromList -List "Final Documents" -ContentType "Document" -Web $web

#This will remove duplicate fields
Remove-PnPField -List "Documents" -Identity "Update ADB Country Document Type" -Force -Web $web
Remove-PnPField -List "Documents" -Identity "Update ADB Document Type" -Force -Web $web
Remove-PnPField -List "Documents" -Identity "Update ADB Project Document Type" -Force -Web $web

echo "END: APPLY CONTENTTYPES"

echo "START: ADDING CONTENT GROUP"

$listToUpdate = @('Documents','Final Documents','Team Tasks','Calendar')
$contentGroupField=$web.Fields.GetByInternalNameOrTitle("ADBContentGroup")
$context.Load($contentGroupField)
$context.ExecuteQuery()

Add-PnPField -List "SitePages" -Field $contentGroupField

foreach($list in $listToUpdate) {
   
    $li=$context.Web.Lists.GetByTitle($list)
    $context.Load($li)
    $context.ExecuteQuery()
    $context.Load($li.ContentTypes)
    $context.ExecuteQuery()

    foreach($ct in $li.ContentTypes) 
    {
        echo $ct.Name 
        if($ct.Name -eq 'ADB Document' -or $ct.Name -eq 'ADB Project Document' -or $ct.Name -eq 'ADB Country Document' -or  $ct.Name -eq 'Task' -or $ct.Name -eq 'Event')
        {

        $link = new-object Microsoft.SharePoint.Client.FieldLinkCreationInformation
        $link.Field = $contentGroupField
        $ct.FieldLinks.Add($link)
        $ct.Update($false)

            try{
                $context.ExecuteQuery()
            } 
            catch
            {
                #swallow unknown error
            }
        }

    }
    
}

echo "END: ADDING CONTENT GROUP"

echo "START: HIDING Content Group FROM Final Docs"

$finalDocs = Get-PnPList -Identity "FinalDocs"
$context.Load($finalDocs)
$context.ExecuteQuery()

$contentTypes = $finalDocs.ContentTypes
$context.Load($contentTypes)
$context.ExecuteQuery()

    
foreach($ct in $contentTypes){        
           # echo $ct.Name      
           if($ct.Name -eq 'ADB Document' -Or $ct.Name -eq 'ADB Country Document' -Or $ct.Name -eq 'ADB Project Document') {
               #load field reference
               $fields = $ct.FieldLinks
               $context.Load($fields)
               $context.ExecuteQuery()
               foreach($ff in $fields) 
               {
                   if(
                   $ff.Name -eq 'ADBDocumentTypeValue' -or 
                   $ff.Name -eq 'ADBContentGroup'
                   ) 
                   {
                       $ff.Hidden = $True
                   }

                   if(
                   $ff.Name -eq 'Title' -or 
                   $ff.Name -eq 'ADBAuthors' -or 
                   $ff.Name -eq 'ADBDepartmentOwner'-or 
                   $ff.Name -eq 'ADBDocumentSecurity'-or 
                   $ff.Name -eq 'ADBDocumentLanguage'-or 
                   $ff.Name -eq 'ADBSourceLink'-or 
                   $ff.Name -eq 'ADBCirculatedLink'               
                   )                   
                   {
                       $ff.Required = $True
                   }
               }
               $ct.Update($false)
               $context.ExecuteQuery()
           }

           if($ct.Name -eq 'ADB Document') {
               #load field reference
               $fields = $ct.FieldLinks
               $context.Load($fields)
               $context.ExecuteQuery()
               foreach($ff in $fields) 
               {
                   if($ff.Name -eq 'ADBDocumentType')                   
                   {
                       $ff.Required = $True
                   }
               }
               $ct.Update($false)
               $context.ExecuteQuery()
           }

           if($ct.Name -eq 'ADB Country Document') {
               #load field reference
               $fields = $ct.FieldLinks
               $context.Load($fields)
               $context.ExecuteQuery()
               foreach($ff in $fields) 
               {
                   if($ff.Name -eq 'ADBCountryDocumentType' -or $ff.Name -eq 'ADBCountry')                   
                   {
                       $ff.Required = $True
                   }
               }
               $ct.Update($false)
               $context.ExecuteQuery()
           }

            if($ct.Name -eq 'ADB Project Document') {
               #load field reference
               $fields = $ct.FieldLinks
               $context.Load($fields)
               $context.ExecuteQuery()
               foreach($ff in $fields) 
               {
                   if($ff.Name -eq 'ADBProjectDocumentType' -or $ff.Name -eq 'ADBCountry' -or $ff.Name -eq 'ADBSector' -or $ff.Name -eq 'ADBProject')             
                   {
                       $ff.Required = $True
                   }
               }
               $ct.Update($false)
               $context.ExecuteQuery()
           }
}

echo "END: HIDING Content Group FROM Final Docs"

echo "START: HIDING Content Group FROM Documents"

$documents = Get-PnPList -Identity "Documents"
$context.Load($documents)
$context.ExecuteQuery()

$contentTypes = $documents.ContentTypes
$context.Load($contentTypes)
$context.ExecuteQuery()

foreach($ct in $contentTypes){        
           # echo $ct.Name      
           if($ct.Name -eq 'ADB Document' -Or $ct.Name -eq 'ADB Country Document' -Or $ct.Name -eq 'ADB Project Document') {
               #load field reference
               $fields = $ct.FieldLinks
               $context.Load($fields)
               $context.ExecuteQuery()
               foreach($ff in $fields) 
               {
                   if(
                   $ff.Name -eq 'ADBDocumentTypeValue' -or 
                   $ff.Name -eq 'ADBContentGroup'
                   ) 
                   {
                       $ff.Hidden = $True
                   }
               }
               $ct.Update($false)
               $context.ExecuteQuery()
           }

           
}

echo "END: HIDING Content Group FROM Documents"


echo "START: HIDING Content Group FROM Team Tasks"

$teamTasks = Get-PnPList -Identity "Team Tasks"
$context.Load($teamTasks)
$context.ExecuteQuery()

$contentTypes = $teamTasks.ContentTypes
$context.Load($contentTypes)
$context.ExecuteQuery()

foreach($ct in $contentTypes){        
           # echo $ct.Name      
           if($ct.Name -eq 'Task') {
               #load field reference
               $fields = $ct.FieldLinks
               $context.Load($fields)
               $context.ExecuteQuery()
               foreach($ff in $fields) 
               {
                   if(
                   $ff.Name -eq 'ADBContentGroup'
                   ) 
                   {
                       $ff.Hidden = $True
                   }
               }
               $ct.Update($false)
               $context.ExecuteQuery()
           }

           
}

echo "END: HIDING Content Group FROM Team Tasks"

echo "START: HIDING Content Group FROM Calendar"

$calendar = Get-PnPList -Identity "Calendar"
$context.Load($calendar)
$context.ExecuteQuery()

$contentTypes = $calendar.ContentTypes
$context.Load($contentTypes)
$context.ExecuteQuery()

foreach($ct in $contentTypes){        
           # echo $ct.Name      
           if($ct.Name -eq 'Event') {
               #load field reference
               $fields = $ct.FieldLinks
               $context.Load($fields)
               $context.ExecuteQuery()
               foreach($ff in $fields) 
               {
                   if(
                   $ff.Name -eq 'ADBContentGroup'
                   ) 
                   {
                       $ff.Hidden = $True
                   }
               }
               $ct.Update($false)
               $context.ExecuteQuery()
           }

           
}

echo "END: HIDING Content Group FROM Calendar"

echo "START: HIDING Content Group FROM Site Pages"

#$sPages = Get-PnPList -Identity "SitePages"

$ctx = Get-PnPContext
$web2 = $ctx.Web
$ctx.Load($web2)
$ctx.Load($web2.Lists)
$sPages = $web2.Lists.GetByTitle("Site Pages")
$spContentTypes = $sPages.ContentTypes
$ctx.Load($sPages)
$ctx.Load($spContentTypes)
$ctx.ExecuteQuery()



foreach($ct2 in $spContentTypes){        
           # echo $ct.Name      
           if($ct2.Name -eq 'Wiki Page') {
               #load field reference
               $fields = $ct2.FieldLinks
               $ctx.Load($fields)
               $ctx.ExecuteQuery()
               foreach($fl in $fields) 
               {
                   if($fl.Name -eq 'ADBContentGroup') 
                   {
                        $fl.DeleteObject();
                        $ct2.Update($false)
                        $ctx.ExecuteQuery()
                        break
                   }
               }
               
           }
}

echo "END: HIDING Content Group FROM Site Pages"



echo "START: APPLY NAVIGATION"
#Apply the Navigation to the target website
Apply-PnPProvisioningTemplate -Path "PNP\collabNAV.xml" -ClearNavigation -Handlers Navigation -Web $web
echo "END: APPLY NAVIGATION"

echo "START: APPLY HOMEPAGE"
#Delete existing homepage
Remove-PnPFile -SiteRelativeUrl "SitePages/Home.aspx" -Force -Web $web

#Apply the Pages template
Apply-PnPProvisioningTemplate -Path "PNP\collabPP.xml" -Handlers  ComposedLook, Pages, PageContents -Web $web

#adds the News webpart to the homepage
$page = Get-PnPClientSidePage -Identity "Home.aspx" -Web $web
Add-PnPClientSideWebPart -Page $page -DefaultWebPartType NewsFeed -Section 2 -Column 1 -Web $web
echo "END: APPLY HOMEPAGE"

echo "START: ENABLE NOSCRIPT"
#Set the NoScriptSite flag to true
Set-PnPTenantSite -Url $targetWebUrl -NoScriptSite:$true
echo "END: ENABLE NOSCRIPT"

echo "END HOMEPAGE"
#>
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
$context.Load($wfh)
$context.Load($wft)
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
}
$context.ExecuteQuery()

echo "END: UPDATE WORKFLOW REFERENCES"

echo "START: UPDATE PERMISSIONS"

$context.Load($web.RoleDefinitions)
$context.ExecuteQuery()

foreach($rd in $web.RoleDefinitions){ 
    if($rd.Name -eq "Edit") 
    {
        $oldBp = $rd.BasePermissions;
        $oldBp.Clear([Microsoft.SharePoint.Client.PermissionKind]::CreateSSCSite)
        $rd.BasePermissions = New-Object Microsoft.SharePoint.Client.BasePermissions

        $rd.BasePermissions = $oldBp
        $rd.Update()
    }

     if($rd.Name -eq "Contribute") 
    {
        $oldBp = $rd.BasePermissions;
        $oldBp.Clear([Microsoft.SharePoint.Client.PermissionKind]::CreateSSCSite)
        $oldBp.Clear([Microsoft.SharePoint.Client.PermissionKind]::DeleteListItems)
        $oldBp.Clear([Microsoft.SharePoint.Client.PermissionKind]::DeleteVersions)
        $rd.BasePermissions = New-Object Microsoft.SharePoint.Client.BasePermissions

        $rd.BasePermissions = $oldBp
        $rd.Update()
    }


    $context.ExecuteQuery()
}


echo "END: UPDATE PERMISSIONS"

echo "START: REMOVE HOME BANNER"
$retval = Get-PnPFile -Url "SitePages/Home.aspx" -AsListItem -Web $web

$x = Set-PnPListItem -List "SitePages" -Identity $retval.Id -Values @{"PageLayoutType"="Home"} -Web $web
echo "END: REMOVE HOME BANNER"

#disconnect
Disconnect-PnPOnline
#exit script
Exit

