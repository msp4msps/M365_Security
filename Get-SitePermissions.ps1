param ( 
    [Parameter(Mandatory = $true)]
    [string] $SiteURL
)

$ReportFile = "C:\Temp\SitePermissionRpt.csv"

Function Get-PnPPermissions([Microsoft.SharePoint.Client.SecurableObject]$Object) {
    $ObjectType = "Site"
    $ObjectTitle = $Object.Title
    $ObjectURL = $Object.Url

    Get-PnPProperty -ClientObject $Object -Property HasUniqueRoleAssignments, RoleAssignments
    $HasUniquePermissions = $Object.HasUniqueRoleAssignments
    $PermissionCollection = @()

    foreach ($RoleAssignment in $Object.RoleAssignments) {
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member

        $PermissionType = $RoleAssignment.Member.PrincipalType
        $LoginName = $RoleAssignment.Member.LoginName.ToLower()
        $PermissionLevels = $RoleAssignment.RoleDefinitionBindings | Select -ExpandProperty Name
        $PermissionLevels = ($PermissionLevels | Where { $_ -ne "Limited Access" -and $_ -ne "Web-Only Limited Access" }) -join ","
        if ([string]::IsNullOrWhiteSpace($PermissionLevels)) { continue }

        # Handle Everyone Except External Users or similar claims
        if ($LoginName -like "*allusers*" -or $LoginName -like "*everyone*") {
            $Permissions = [PSCustomObject]@{
                Object              = $ObjectType
                Title               = $ObjectTitle
                URL                 = $ObjectURL
                HasUniquePermissions = $HasUniquePermissions
                Users               = "Everyone Except External Users"
                UserPrincipalName   = "Built-in"
                Type                = "Claim (Everyone)"
                Permissions         = $PermissionLevels
                GrantedThrough      = "Built-in Group"
            }
            $PermissionCollection += $Permissions
            continue
        }

        # SharePoint Group Handling
        if ($PermissionType -eq "SharePointGroup") {
            $GroupMembers = Get-PnPGroupMember -Identity $RoleAssignment.Member.LoginName
            if ($GroupMembers.Count -eq 0) { continue }

            foreach ($GroupMember in $GroupMembers) {
                $isHiddenInUI = Get-PnPProperty -ClientObject $GroupMember -Property IsHiddenInUI
                if ($isHiddenInUI -eq $true) { continue }

                $UPN = Get-PnPProperty -ClientObject $GroupMember -Property UserPrincipalName
                $Permissions = [PSCustomObject]@{
                    Object              = $ObjectType
                    Title               = $ObjectTitle
                    URL                 = $ObjectURL
                    HasUniquePermissions = $HasUniquePermissions
                    Users               = $GroupMember.Title
                    UserPrincipalName   = $UPN
                    Type                = $PermissionType
                    Permissions         = $PermissionLevels
                    GrantedThrough      = "SharePoint Group: $($RoleAssignment.Member.LoginName)"
                }
                $PermissionCollection += $Permissions
            }
        }
        else {
            # Direct User
            $UPN = Get-PnPProperty -ClientObject $RoleAssignment.Member -Property UserPrincipalName
            $Permissions = [PSCustomObject]@{
                Object              = $ObjectType
                Title               = $ObjectTitle
                URL                 = $ObjectURL
                HasUniquePermissions = $HasUniquePermissions
                Users               = $RoleAssignment.Member.Title
                UserPrincipalName   = $UPN
                Type                = $PermissionType
                Permissions         = $PermissionLevels
                GrantedThrough      = "Direct Permissions"
            }
            $PermissionCollection += $Permissions
        } 
    }

    #Check if the site (or its objects) contains any Direct permissions to "Everyone except external users"
    $EEEUsers = Get-PnPUser  | Where {$_.Title -eq "Everyone except external users"}
    If($EEEUsers){
        $Permissions = [PSCustomObject]@{
            Object              = $ObjectType
            Title               = $ObjectTitle
            URL                 = $ObjectURL
            HasUniquePermissions = $HasUniquePermissions
            Users               = "Everyone Except External Users"
            UserPrincipalName   = "N/A"
            Type                = 'Security Group'
            Permissions         = 'Edit'
            GrantedThrough      = "Direct Permissions"
        }
        $PermissionCollection += $Permissions
    }

    # Get M365 Group Members and Owners
    $siteGroup = Get-PnPMicrosoft365Group -Identity $ObjectTitle -ErrorAction SilentlyContinue
    if ($siteGroup) {
        $owners = Get-PnPMicrosoft365GroupOwners -Identity $siteGroup.Id
        $members = Get-PnPMicrosoft365GroupMembers -Identity $siteGroup.Id

        foreach ($owner in $owners) {
            $Permissions = [PSCustomObject]@{
                Object              = $ObjectType
                Title               = $ObjectTitle
                URL                 = $ObjectURL
                HasUniquePermissions = $HasUniquePermissions
                Users               = $owner.DisplayName
                UserPrincipalName   = $owner.UserPrincipalName
                Type                = "M365 Group Owner"
                Permissions         = "Group Owner"
                GrantedThrough      = "Microsoft 365 Group"
            }
            $PermissionCollection += $Permissions
        }

        foreach ($member in $members) {
            $Permissions = [PSCustomObject]@{
                Object              = $ObjectType
                Title               = $ObjectTitle
                URL                 = $ObjectURL
                HasUniquePermissions = $HasUniquePermissions
                Users               = $member.DisplayName
                UserPrincipalName   = $member.UserPrincipalName
                Type                = "M365 Group Member"
                Permissions         = "Group Member"
                GrantedThrough      = "Microsoft 365 Group"
            }
            $PermissionCollection += $Permissions
        }
    }

    $PermissionCollection | Export-Csv -Path $ReportFile -NoTypeInformation -Append
}

Function Generate-PnPSitePermissionRpt {
    [cmdletbinding()]
    Param (
        [Parameter(Mandatory = $true)]
        [String] $SiteURL,
        [Parameter(Mandatory = $true)]
        [String] $ReportFile
    )
    Try {
        Connect-PnPOnline -Url $SiteURL -Interactive

        $Web = Get-PnPWeb

        # Site Collection Admins
        $SiteAdmins = Get-PnPSiteCollectionAdmin
        $SiteCollectionAdmins = ($SiteAdmins | Select -ExpandProperty Title) -join ","

        $Permissions = [PSCustomObject]@{
            Object              = "Site Collection"
            Title               = $Web.Title
            URL                 = $Web.Url
            HasUniquePermissions = $true
            Users               = $SiteCollectionAdmins
            UserPrincipalName   = "N/A"
            Type                = "Site Collection Administrators"
            Permissions         = "Site Owner"
            GrantedThrough      = "Direct Permissions"
        }

        $Permissions | Export-Csv -Path $ReportFile -NoTypeInformation

        # Get site-level permissions only
        Get-PnPPermissions -Object $Web

        Write-host -f Green "`n*** Site-Level Permission Report Generated Successfully! ***"
    }
    Catch {
        Write-host -f Red "Error Generating Site Permission Report!" $_.Exception.Message
    }
}

# Run the report
Generate-PnPSitePermissionRpt -SiteURL $SiteURL -ReportFile $ReportFile