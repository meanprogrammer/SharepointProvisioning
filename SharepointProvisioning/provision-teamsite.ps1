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

#Sets the field status if hidden or required
Function SetFieldStatus
{
    Param(
        [bool]$hidden = $false,
        [bool]$required = $false,
        [string]$listName,
        [string[]]$contentTypes,
        [string[]]$fields,
        [Parameter(Mandatory=$True)]
        [Microsoft.SharePoint.Client.ClientContext]$context
    )
    #load the list
    $list = Get-PnPList -Identity $listName
    $context.Load($list)
    $context.ExecuteQuery()
    #load the list content types    
    $cts = $list.ContentTypes
    $context.Load($cts)
    $context.ExecuteQuery()

    foreach($ct in $cts){        
            #checks if the content types exist in the $contentTypes array
            if($contentTypes -contains $ct.Name) {
                #load field reference
                $flinks = $ct.FieldLinks
                $context.Load($flinks)
                $context.ExecuteQuery()
                #for each field links, if it exists on the fields array
                #set the hidden and the required
                foreach($ff in $flinks) 
                {
                    if($fields -contains $ff.Name){ 
                        $ff.Hidden = $hidden
                        $ff.Required = $Required
                    }
                }
                #update the content type
                $ct.Update($false)
                $context.ExecuteQuery()
            }         
    }

    
}

#adds the content group field to a list content type
Function AddContentGroup
{
    Param(
        [string]$listName,
        [string[]]$contentTypes,
        [string[]]$fields,
        [Parameter(Mandatory=$True)]
        [Microsoft.SharePoint.Client.ClientContext]$context
    )   
    #load the list
    $li=$context.Web.Lists.GetByTitle($listName)
    $context.Load($li)
    $context.ExecuteQuery()
    #load the list content types
    $context.Load($li.ContentTypes)
    $context.ExecuteQuery()

    foreach($ct in $li.ContentTypes) 
    {
        #checks if the content types exist in the $contentTypes array
        if($contentTypes -contains $ct.Name) 
        {
            #create instance of FieldLinkCreationInformation
            $link = new-object Microsoft.SharePoint.Client.FieldLinkCreationInformation
            #set field to content group
            $link.Field = $contentGroupField
            #add to field links
            $ct.FieldLinks.Add($link)
            #update the content type
            $ct.Update($false)

            try{
                $context.ExecuteQuery()
            } 
            catch
            {
                $ErrorMessage = $_.Exception.Message
                $FailedItem = $_.Exception.ItemName
                echo "ERROR: $ErrorMessage $FailedItem"
                #swallow unknown error
            }
        }

    }
}

Function RemoveFieldLink
{
    Param(
        [string]$listName,
        [string]$contentType,
        [string]$field,
        [Parameter(Mandatory=$True)]
        [Microsoft.SharePoint.Client.ClientContext]$context,
        [Parameter(Mandatory=$True)]
        [Microsoft.SharePoint.Client.Web]$web
    )

    $sPages = $web.Lists.GetByTitle($listName)
    $spContentTypes = $sPages.ContentTypes
    $context.Load($sPages)
    $context.Load($spContentTypes)
    $context.ExecuteQuery()

    foreach($ct in $spContentTypes){        
           # echo $ct.Name      
           if($ct.Name -eq $contentType) {
               #load field reference
               $fields = $ct.FieldLinks
               $context.Load($fields)
               $context.ExecuteQuery()
               foreach($fl in $fields) 
               {
                   if($fl.Name -eq $field) 
                   {
                        $fl.DeleteObject();
                        $ct.Update($false)
                        $context.ExecuteQuery()
                        break
                   }
               }
               
           }
    }   
}

Function UpdateWorkflowReferences
{
    Param(
        [string]$listName,
        [string]$workflowHistory,
        [string]$workflowTask,
        [string[]]$contentTypes,
        [Parameter(Mandatory=$True)]
        [Microsoft.SharePoint.Client.ClientContext]$context
    )

    #Gets the site of the target website
    $site = $context.Site
    $context.Load($site)
    $context.ExecuteQuery()

    #gets reference to the "Documents" document library
    $list = $web.Lists.GetByTitle($listName)

    $context.Load($list)
    $context.ExecuteQuery()

    #loads all workflow associations
    $context.Load($list.WorkflowAssociations)
    $context.ExecuteQuery()  

    #Gets the WorkflowServicesManager instance
    $servicesManager = New-Object Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager($context, $web)
    #Gets the WorkflowSubscriptionService
    $subscriptionService = $servicesManager.GetWorkflowSubscriptionService()
    #List all the subscription
    $subscriptions = $subscriptionService.EnumerateSubscriptionsByList($list.Id)

    $context.Load($subscriptions)
    $context.ExecuteQuery()

    #Gets a reference to the Workflow history list
    $wfh = $web.Lists.GetByTitle($workflowHistory)
    $wft = $web.Lists.GetByTitle($workflowTask)

    #Gets a reference to the Workflow history task
    $context.Load($wfh)
    $context.Load($wft)
    $context.ExecuteQuery()

    #Loop through all the subscription and
    #set the HistoryListId and TaskListId
    #and publish
    foreach ($s in $subscriptions)
    {
        echo $s.Name
        if ($contentTypes -contains $s.Name)
        {
            
            $s.SetProperty("HistoryListId", $wfh.Id)
            $s.SetProperty("TaskListId", $wft.Id)
            $s.SetProperty("FormData", "")
            $subscriptionService.PublishSubscriptionForList($s, $list.Id)
        } 
    }

    $context.ExecuteQuery()
}

#removes permission from role definition
Function RemoveFromPermission
{
    Param(
        [string]$role,
        [Microsoft.SharePoint.Client.PermissionKind[]]$permissionToRemove, 
        [Parameter(Mandatory=$True)]
        [Microsoft.SharePoint.Client.Web]$web,
        [Parameter(Mandatory=$True)]
        [Microsoft.SharePoint.Client.ClientContext]$context
    )

    #loads role definitions
    $context.Load($web.RoleDefinitions)
    $context.ExecuteQuery()

    foreach($rd in $web.RoleDefinitions){ 
        if($rd.Name -eq $role) 
        {
            $oldBp = $rd.BasePermissions
            #remove each permission that is in the array
            foreach ($permission in $permissionToRemove) {
                $oldBp.Clear($permission)
            }
            $rd.BasePermissions = New-Object Microsoft.SharePoint.Client.BasePermissions
            #assign the modified permission to the role definition BasePermissions
            $rd.BasePermissions = $oldBp
            $rd.Update()
        }
        $context.ExecuteQuery()
    }
}

#End Functions

echo "BEGIN SITE PROVISIONING"

#Set-PnPTraceLog -On -Level Debug
Set-PnPTraceLog -On -LogFile traceoutput.txt -Level Error

#Set variables
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

$contentGroupField=$web.Fields.GetByInternalNameOrTitle("ADBContentGroup")
$context.Load($contentGroupField)
$context.ExecuteQuery()

Add-PnPField -List "SitePages" -Field $contentGroupField

AddContentGroup -listName "Documents" -contentTypes @('Document','ADB Document','ADB Project Document','ADB Country Document','Task','Event') -context $context
AddContentGroup -listName "Final Documents" -contentTypes @('ADB Document','ADB Project Document','ADB Country Document','Task','Event') -context $context
AddContentGroup -listName "Team Tasks" -contentTypes @('Task') -context $context
AddContentGroup -listName "Calendar" -contentTypes @('Event') -context $context

echo "END: ADDING CONTENT GROUP"

echo "START: HIDING Content Group FROM Final Docs"
SetFieldStatus -hidden 1 -listName "FinalDocs" -contentTypes @('ADB Document','ADB Country Document','ADB Project Document') -fields @('ADBDocumentTypeValue','ADBContentGroup') -context $context
SetFieldStatus -required 1 -listName "FinalDocs" -contentTypes @('ADB Document','ADB Country Document','ADB Project Document') -fields @('Title','ADBAuthors','ADBDepartmentOwner','ADBDocumentSecurity','ADBDocumentLanguage','ADBSourceLink','ADBCirculatedLink') -context $context
SetFieldStatus -required 1 -listName "FinalDocs" -contentTypes @('ADB Document') -fields @('ADBDocumentType') -context $context
SetFieldStatus -required 1 -listName "FinalDocs" -contentTypes @('ADB Country Document') -fields @('ADBCountryDocumentType','ADBCountry') -context $context
SetFieldStatus -required 1 -listName "FinalDocs" -contentTypes @('ADB Project Document') -fields @('ADBProjectDocumentType','ADBCountry','ADBSector','ADBProject') -context $context
echo "END: HIDING Content Group FROM Final Docs"

echo "START: HIDING Content Group FROM Documents"
SetFieldStatus -hidden 1 -listName "Documents" -contentTypes @('Document','ADB Document','ADB Country Document','ADB Project Document') -fields @('ADBDocumentTypeValue','ADBContentGroup') -context $context
echo "END: HIDING Content Group FROM Documents"

echo "START: HIDING Content Group FROM Team Tasks"
SetFieldStatus -hidden 1 -listName "Team Tasks" -contentTypes @('Task') -fields @('ADBContentGroup') -context $context
echo "END: HIDING Content Group FROM Team Tasks"

echo "START: HIDING Content Group FROM Calendar"
SetFieldStatus -hidden 1 -listName "Calendar" -contentTypes @('Event') -fields @('ADBContentGroup') -context $context
echo "END: HIDING Content Group FROM Calendar"

echo "START: HIDING Content Group FROM Site Pages"
RemoveFieldLink -listName "Site Pages" -contentType "Wiki Page" -field "ADBContentGroup" -context $context -web $web
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

#Updates the workflow references
echo "START: UPDATE WORKFLOW REFERENCES"
#This is required to use the WorkflowServicesManager class
Add-Type -Path "Microsoft.SharePoint.Client.WorkflowServices.dll"
UpdateWorkflowReferences -listName "Documents" -workflowHistory "Workflow History" -workflowTask "Update Document Type Workflow Tasks" -contentTypes @('Update ADB Project Document Type','Update ADB Country Document Type','Update ADB Document Type') -context $context

echo "END: UPDATE WORKFLOW REFERENCES"

echo "START: UPDATE PERMISSIONS"
RemoveFromPermission -role "Edit" -permissionToRemove @([Microsoft.SharePoint.Client.PermissionKind]::CreateSSCSite) -web $web -context $context
RemoveFromPermission -role "Contribute" -permissionToRemove @([Microsoft.SharePoint.Client.PermissionKind]::CreateSSCSite,[Microsoft.SharePoint.Client.PermissionKind]::DeleteListItems,[Microsoft.SharePoint.Client.PermissionKind]::DeleteVersions) -web $web -context $context

echo "END: UPDATE PERMISSIONS"

echo "START: REMOVE HOME BANNER"
$retval = Get-PnPFile -Url "SitePages/Home.aspx" -AsListItem -Web $web

$x = Set-PnPListItem -List "SitePages" -Identity $retval.Id -Values @{"PageLayoutType"="Home"} -Web $web
echo "END: REMOVE HOME BANNER"

#disconnect
Disconnect-PnPOnline
#exit script
Exit

