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

Function CheckIfSiteExists()
{
    Param([string]$url)
    Try
    {
        #echo $url
        $existingSite = Get-PnPTenantSite -url $url -Detailed -ErrorAction Stop
        Start-Sleep -Seconds 5
        if($existingSite -ne $null)
        {
            echo "Target site already exist. Terminating provisioning script."
            Exit
        }
    }
    catch
    {
        #swallowing exception
        echo $_.Exception.Message
    }

}

#End Functions