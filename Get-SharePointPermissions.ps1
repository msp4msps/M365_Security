param ( 
    [Parameter(Mandatory=$false)]
    [string] $SiteURL
)

#Function to Get Permissions Applied on a particular Object, such as: Web, List, Folder or List Item
Function Get-PnPPermissions([Microsoft.SharePoint.Client.SecurableObject]$Object)
{
    #Determine the type of the object
    Switch($Object.TypedObject.ToString())
    {
        "Microsoft.SharePoint.Client.Web"  { $ObjectType = "Site" ; $ObjectURL = $Object.URL; $ObjectTitle = $Object.Title }
        "Microsoft.SharePoint.Client.ListItem"
        {
            Write-Host "Getting data for $($Object.FieldValues.FileLeafRef)"
            If($Object.FileSystemObjectType -eq "Folder")
            {
                $ObjectType = "Folder"
                $ObjectTitle = $Object.FieldValues.FileLeafRef
                $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''),$Object.FieldValues.FileRef)
            }
            If($Object.FileSystemObjectType -eq "File")
            {
                $ObjectType = "File"
                $ObjectTitle = $Object.FieldValues.FileLeafRef
                $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''),$Object.FieldValues.FileRef)
            }
            If($Object.FileSystemObjectType -eq "ListItem")
            {
                $ObjectType = "List Item"
                $ObjectTitle = $Object["Title"]
                #Get the URL of the List Item
                $DefaultDisplayFormUrl = Get-PnPProperty -ClientObject $Object.ParentList -Property DefaultDisplayFormUrl                    
                $ObjectURL = $("{0}{1}?ID={2}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''), $DefaultDisplayFormUrl,$Object.ID)
            }
        }
        Default
        {
            $ObjectType = "List or Library"
            $ObjectTitle = $Object.Title
            #Get the URL of the List or Library
            $RootFolder = Get-PnPProperty -ClientObject $Object -Property RootFolder    
            $ObjectURL = $("{0}{1}" -f $Web.Url.Replace($Web.ServerRelativeUrl,''), $RootFolder.ServerRelativeUrl)
        }
    }
   
    #Get permissions assigned to the object
    Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments
 
    #Check if Object has unique permissions
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments
    $ShareLinkVariable = $null
  
     
    #Loop through each permission assigned and extract details
    $PermissionCollection = @()
    Foreach($RoleAssignment in $Object.RoleAssignments)
    {
        #Get the Permission Levels assigned and Member
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
        # $test = $RoleAssignment.Member| Select-Object -Property *
        # #
 
        #Get the Principal Type: User, SP Group, AD Group
        $PermissionType = $RoleAssignment.Member.PrincipalType

        #Get the Permission Levels assigned
        $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select -ExpandProperty Name
        if($PermissionLevels -eq "Web-Only Limited Access") {Continue}
 
        #Remove Limited Access
        $PermissionLevels = ($PermissionLevels | Where { $_ -ne "Limited Access"}) -join ","
 
        #Leave Principals with no Permissions
        If($PermissionLevels.Length -eq 0) {Continue}
 
        #Get SharePoint group members
        If($PermissionType -eq "SharePointGroup")
        {
            #Get Group Members
            $GroupMembers = Get-PnPGroupMember -Identity $RoleAssignment.Member.LoginName
            
            #Define Granted through if text starts with SharingLinks
            if($RoleAssignment.Member.LoginName -like "*SharingLinks*") {
                $GrantedThrough = "Sharing Link"      
                        
            } else {
                $GrantedThrough = "SharePoint Group: $($RoleAssignment.Member.LoginName)"
            }
                 
            #Leave Empty Groups
            If($GroupMembers.count -eq 0){Continue}
            
            #Instead of bundling users, create a separate row for each group member
            Foreach($GroupMember in $GroupMembers)
            {  
                
                $isHiddeninUI = Get-PnPProperty -ClientObject $GroupMember -Property IsHiddenInUI
                If($isHiddeninUI -eq $true) {Continue}
                $UPN = Get-PnPProperty -ClientObject $GroupMember -Property UserPrincipalName
                # If($UPN -eq $null) {Continue}
                #Add the Data to Object for each user
                $Permissions = New-Object PSObject
                $Permissions | Add-Member NoteProperty Object($ObjectType)
                $Permissions | Add-Member NoteProperty Title($ObjectTitle)
                $Permissions | Add-Member NoteProperty URL($ObjectURL)
                $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
                $Permissions | Add-Member NoteProperty Users($GroupMember.Title)
                $UserPrincipalName = Get-PnPProperty -ClientObject $GroupMember -Property UserPrincipalName
                $Permissions | Add-Member NoteProperty UserPrincipalName($UserPrincipalName)
                $Permissions | Add-Member NoteProperty Type($PermissionType)
                $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
                $Permissions | Add-Member NoteProperty GrantedThrough($GrantedThrough)
                $PermissionCollection += $Permissions
            }
        }
        Else
        {
            #Add the Data to Object for direct users
            $Permissions = New-Object PSObject
            $Permissions | Add-Member NoteProperty Object($ObjectType)
            $Permissions | Add-Member NoteProperty Title($ObjectTitle)
            $Permissions | Add-Member NoteProperty URL($ObjectURL)
            $Permissions | Add-Member NoteProperty HasUniquePermissions($HasUniquePermissions)
            $Permissions | Add-Member NoteProperty Users($RoleAssignment.Member.Title)
            #get user email
            $UserPrincipalName = Get-PnPProperty -ClientObject $RoleAssignment.Member -Property UserPrincipalName
            $Permissions | Add-Member NoteProperty UserPrincipalName($UserPrincipalName)
            $Permissions | Add-Member NoteProperty Type($PermissionType)
            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
            $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
            $PermissionCollection += $Permissions
        }
    }
    #Export Permissions to CSV File
    $PermissionCollection | Export-CSV $ReportFile -NoTypeInformation -Append
}

   
#Function to get sharepoint online site permissions report
Function Generate-PnPSitePermissionRpt()
{
[cmdletbinding()]
 
    Param 
    (   
        [Parameter(Mandatory=$false)] [String] $SiteURL,
        [Parameter(Mandatory=$false)] [String] $ReportFile,        
        [Parameter(Mandatory=$false)] [switch] $Recursive,
        [Parameter(Mandatory=$false)] [switch] $ScanItemLevel,
        [Parameter(Mandatory=$false)] [switch] $IncludeInheritedPermissions       
    ) 
    Try {
        #Connect to the Site
        Connect-PnPOnline -URL $SiteURL -UseWebLogin
        #Get the Web
        $Web = Get-PnPWeb
 
        Write-host -f Yellow "Getting Site Collection Administrators..."
        #Get Site Collection Administrators
        $SiteAdmins = Get-PnPSiteCollectionAdmin
         
        $SiteCollectionAdmins = ($SiteAdmins | Select -ExpandProperty Title) -join ","
        #Add the Data to Object
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty Object("Site Collection")
        $Permissions | Add-Member NoteProperty Title($Web.Title)
        $Permissions | Add-Member NoteProperty URL($Web.URL)
        $Permissions | Add-Member NoteProperty HasUniquePermissions("TRUE")
        $Permissions | Add-Member NoteProperty Users($SiteCollectionAdmins)
        $Permissions | Add-Member NoteProperty UserPrincipalName("N/A")
        $Permissions | Add-Member NoteProperty Type("Site Collection Administrators")
        $Permissions | Add-Member NoteProperty Permissions("Site Owner")
        $Permissions | Add-Member NoteProperty GrantedThrough("Direct Permissions")
               
        #Export Permissions to CSV File
        $Permissions | Export-CSV $ReportFile -NoTypeInformation
   
        #Function to Get Permissions of All List Items of a given List
        Function Get-PnPListItemsPermission([Microsoft.SharePoint.Client.List]$List)
        {
            Write-host -f Yellow "`t `t Getting Permissions of List Items in the List:"$List.Title
  
            #Get All Items from List in batches
            $ListItems = Get-PnPListItem -List $List -PageSize 500
  
            $ItemCounter = 0
            #Loop through each List item
            ForEach($ListItem in $ListItems)
            {
                #Get Objects with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                If($IncludeInheritedPermissions)
                {
                    Get-PnPPermissions -Object $ListItem
                }
                Else
                {
                    #Check if List Item has unique permissions
                    $HasUniquePermissions = Get-PnPProperty -ClientObject $ListItem -Property HasUniqueRoleAssignments
                    If($HasUniquePermissions -eq $True)
                    {
                        #Call the function to generate Permission report
                        Get-PnPPermissions -Object $ListItem
                    }
                }
                $ItemCounter++
                Write-Progress -PercentComplete ($ItemCounter / ($List.ItemCount) * 100) -Activity "Processing Items $ItemCounter of $($List.ItemCount)" -Status "Searching Unique Permissions in List Items of '$($List.Title)'"
            }
        }
 
        #Function to Get Permissions of all lists from the given web
        Function Get-PnPListPermission([Microsoft.SharePoint.Client.Web]$Web)
        {
            #Get All Lists from the web
            $Lists = Get-PnPProperty -ClientObject $Web -Property Lists
   
            #Exclude system lists
            $ExcludedLists = @("Access Requests","App Packages","appdata","appfiles","Apps in Testing","Cache Profiles","Composed Looks","Content and Structure Reports","Content type publishing error log","Converted Forms",
            "Device Channels","Form Templates","fpdatasources","Get started with Apps for Office and SharePoint","List Template Gallery", "Long Running Operation Status","Maintenance Log Library", "Images", "site collection images"
            ,"Master Docs","Master Page Gallery","MicroFeed","NintexFormXml","Quick Deploy Items","Relationships List","Reusable Content","Reporting Metadata", "Reporting Templates", "Search Config List","Site Assets","Preservation Hold Library",
            "Site Pages", "Solution Gallery","Style Library","Suggested Content Browser Locations","Theme Gallery", "TaxonomyHiddenList","User Information List","Web Part Gallery","wfpub","wfsvc","Workflow History","Workflow Tasks", "Pages")
             
            $Counter = 0
            #Get all lists from the web  
            ForEach($List in $Lists)
            {
                #Exclude System Lists
                If($List.Hidden -eq $False -and $ExcludedLists -notcontains $List.Title)
                {
                    $Counter++
                    Write-Progress -PercentComplete ($Counter / ($Lists.Count) * 100) -Activity "Exporting Permissions from List '$($List.Title)' in $($Web.URL)" -Status "Processing Lists $Counter of $($Lists.Count)"
 
                    #Get Item Level Permissions if 'ScanItemLevel' switch present
                    If($ScanItemLevel)
                    {
                        #Get List Items Permissions
                        Get-PnPListItemsPermission -List $List
                    }
 
                    #Get Lists with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-PnPPermissions -Object $List
                    }
                    Else
                    {
                        #Check if List has unique permissions
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $List -Property HasUniqueRoleAssignments
                        If($HasUniquePermissions -eq $True)
                        {
                            #Call the function to check permissions
                            Get-PnPPermissions -Object $List
                        }
                    }
                }
            }
        }
   
        #Function to Get Webs's Permissions from given URL
        Function Get-PnPWebPermission([Microsoft.SharePoint.Client.Web]$Web)
        {
            #Call the function to Get permissions of the web
            Write-host -f Yellow "Getting Permissions of the Web: $($Web.URL)..." 
            Get-PnPPermissions -Object $Web
   
            #Get List Permissions
            Write-host -f Yellow "`t Getting Permissions of Lists and Libraries..."
            Get-PnPListPermission($Web)
 
            #Recursively get permissions from all sub-webs based on the "Recursive" Switch
            If($Recursive)
            {
                #Get Subwebs of the Web
                $Subwebs = Get-PnPProperty -ClientObject $Web -Property Webs
 
                #Iterate through each subsite in the current web
                Foreach ($Subweb in $web.Webs)
                {
                    #Get Webs with Unique Permissions or Inherited Permissions based on 'IncludeInheritedPermissions' switch
                    If($IncludeInheritedPermissions)
                    {
                        Get-PnPWebPermission($Subweb)
                    }
                    Else
                    {
                        #Check if the Web has unique permissions
                        $HasUniquePermissions = Get-PnPProperty -ClientObject $SubWeb -Property HasUniqueRoleAssignments
   
                        #Get the Web's Permissions
                        If($HasUniquePermissions -eq $true)
                        {
                            #Call the function recursively                           
                            Get-PnPWebPermission($Subweb)
                        }
                    }
                }
            }
        }
 
        #Call the function with RootWeb to get site collection permissions
        Get-PnPWebPermission $Web
   
        Write-host -f Green "`n*** Site Permission Report Generated Successfully!***"
     }
    Catch {
        write-host -f Red "Error Generating Site Permission Report!" $_.Exception.Message
   }
}
   
#region ***Parameters***
$ReportFile="C:\Temp\SitePermissionRpt.csv"
#endregion
 
#Call the function to generate permission report
Generate-PnPSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile -Recursive -ScanItemLevel 
#Generate-PnPSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile -Recursive -ScanItemLevel -IncludeInheritedPermissions