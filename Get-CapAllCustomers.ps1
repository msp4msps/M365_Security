param ( 
    [Parameter(Mandatory=$false)]
    [string] $AppId, 
    [Parameter(Mandatory=$false)]
    [string] $AppSecret, 
    [Parameter(Mandatory=$false)]
    [string] $TenantId,
    [Parameter(Mandatory=$false)]
    [string] $refreshToken
)

   <#
        .SYNOPSIS
            Automatically generate an Excel report containing your current Conditional Access policies across your customers that have a GDAP relationship.

        .DESCRIPTION
            Uses Microsoft Graph to fetch all Conditional Access policies and exports an Excel report.

            To make the report easier to read, do this:
            1. Select all cells.
            2. Click on "Wrap Text".
            3. Click on "Top Align".

        This report requires you have set up the necessary prerequisites for the Secure Applicable Model and GDAP in partner center.
        For more information, check out this blog: https://tminus365.com/my-automations-break-with-gdap-the-fix/

        .PARAMETER
            AppId
            The AppId of the Secure Applicable Model. AKA Client ID of the app registration.

        .PARAMETER
            AppSecret
            The AppSecret of the Secure Applicable Model. Aka the Client Secret on the app registration. 

        .PARAMETER 
            TenantId
            Your Partner Tenant ID

        .PARAMETER 
            refreshToken
            The refreshToken of the Secure Applicable Model.
        
        .Minumum Requirements
            - PowerShell 7.1
            - Microsoft Graph Permissions on App Registration: DelegatedAdmin.ReadWrite.All, Policy.Read.All,Directory.Read.All,Groups.Read.All
            - Min GDAP Permissions: Conditional Access Administrator, Directory Reader


        .INPUTS
            Your AppId, AppSecret, TenantId, and refreshToken from the Secure Applicable Model.

        .OUTPUTS
            Excel report with all Conditional Access policies.

        .NOTES
            Author:   Nick Ross
            GitHub:   https://github.com/msp4msps/M365_Security
            Blog:     https://tminus365.com/
    
    #>


# ----- [Initialisations] -----

    # Set Error Action - Possible choices: Stop, SilentlyContinue
    $ErrorActionPreference = "SilentlyContinue"
    
    
    # ----- [Execution] -----

    #Check if the Excel module is installed.
    if (Get-Module -ListAvailable -Name "ImportExcel") {
        # Do nothing.
    } else {
        Write-Error -Exception "The Excel PowerShell module is not installed. Please, run 'Install-Module ImportExcel -Force' as an admin and try again." -ErrorAction Stop
    }

    #Check to see if Graph module is installed
    if (Get-Module -ListAvailable -Name "Microsoft.Graph") {
        # Do nothing.
    } else {
        #Install the module
        Install-Module -Name "Microsoft.Graph" -Force -AllowClobber
    }

# Initialize a list to hold audit logs
$AuditLogs = [System.Collections.Generic.List[Object]]::new()


#function to record an audit log of events
function Write-Log {
    param (
        [string]$customerName,
        [string]$EventName,
        [string]$Status,
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Result = @{

        'Customer Name' = $CustomerName
        'Timestamp' = $timestamp
        'Event' = $EventName
        'Status' = $Status
        'Message' = $Message
    }  
    $ReportLine = New-Object PSObject -Property $Result
    $AuditLogs.Add($ReportLine)
    

    # Optionally, you can also display this log message in the console
    Write-Host $logMessage
}

# Function to get an access token for Microsoft Graph
function Get-GraphAccessToken ($TenantId) {
    $Body = @{
        'tenant'        = $TenantId
        'client_id'     = $AppId
        'scope'         = 'https://graph.microsoft.com/.default'
        'client_secret' = $AppSecret
        'grant_type'    = 'refresh_token'
        'refresh_token' = $refreshToken
    }
    $Params = @{
        'Uri'         = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        'Method'      = 'Post'
        'ContentType' = 'application/x-www-form-urlencoded'
        'Body'        = $Body
    }
    $AuthResponse = Invoke-RestMethod @Params
    return $AuthResponse.access_token
}


# Get the access token
try {
    $AccessToken =  Get-GraphAccessToken -TenantId $TenantId
    Write-Host "Got partner access token for Microsoft Graph" -ForegroundColor Green
    Write-Log -customerName "Internal Tenant" -EventName "Get-GraphAccessToken" -Status "Success" -Message "Got partner access token for Microsoft Graph"

}
catch {
    Write-Host "Got partner access token for Microsoft Graph" -ForegroundColor Green
    $ErrorMessage = $_.Exception.Message
    Write-Log -customerName "Internal Tenant" -EventName "Get-GraphAccessToken" -Status "Fail" -Message $ErrorMessage
    break
}
# Define header with authorization token
$Headers = @{
    'Authorization' = "Bearer $AccessToken"
    'Content-Type'  = 'application/json'
}

# Get GDAP Customers
try{
    $GraphUrl = "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers"
    $Customers = Invoke-RestMethod -Uri $GraphUrl  -Headers $Headers -Method Get
    #Write-Host a list of the customer display Names
    Write-Host "Customers:" -ForegroundColor Green
    $Customers.value.displayName
    Write-Log -customerName "Internal Tenant" -EventName "Get-GDAPCustomers" -Status "Success" -Message "Got GDAP Customers"

}catch{
    Write-Host "Failed to get GDAP Customers" -ForegroundColor Red
    $ErrorMessage = $_.Exception.Message
    Write-Log -customerName "Internal Tenant" -EventName "Get-GDAPCustomers" -Status "Fail" -Message $ErrorMessage
    break

}

 #Create object to hold all results for customers
 $AllResults = @()

#loop through customers and make sure to continue in try catch block if any errors

$Customers.value | ForEach-Object {
    $Customer = $_
    $CustomerName = $Customer.displayName
    $CustomerTenantId = $Customer.id
    Write-Host "Processing Customer: $CustomerName" -ForegroundColor Blue
   # Get the access token
    try{
        $Token =  Get-GraphAccessToken -TenantId $CustomerTenantId
        $AccessToken = ConvertTo-SecureString $Token -AsPlainText -Force
        Write-Host "Got access token for Microsoft Graph" -ForegroundColor Green
        Write-Log -customerName $CustomerName -EventName "Get-GraphAccessToken" -Status "Success" -Message "Got access token for Microsoft Graph"

    }
    catch{
        Write-Host "Failed to get access token for Microsoft Graph" -ForegroundColor Red
        Write-Host $_.Exception.Message
        Write-Log -customerName $CustomerName -EventName "Get-GraphAccessToken" -Status "Fail" -Message $_.Exception.Message
        return
    }

    try{
        # Connect to Microsoft Graph.
        Connect-MgGraph -AccessToken $AccessToken -NoWelcome
        Write-host "Connected to Microsoft Graph" -ForegroundColor Green
        Write-Log -customerName $CustomerName -EventName "Connect-MgGraph" -Status "Success" -Message "Connected to Microsoft Graph"
    } 
    catch {
        Write-Host "Failed to connect to Microsoft Graph" -ForegroundColor Red
        Write-Host $_.Exception.Message
        Write-Log -customerName $CustomerName -EventName "Connect-MgGraph" -Status "Fail" -Message $_.Exception.Message
        return 
    }

        try{
            $CAPolicies = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/beta/identity/conditionalAccess/policies').value | ConvertTo-Json -Depth 10 | ConvertFrom-Json
            Write-Host "Processing Conditional Access Policies..." -ForegroundColor Green
            Write-Log -customerName $CustomerName -EventName "Get-ConditionalAccessPolicies" -Status "Success" -Message "Got Conditional Access Policies"

            # Fetch service principals for id translation.
            $EnterpriseApps = Get-MgServicePrincipal

            # Fetch roles for id translation.
            $EntraIDRoles =  Get-MgDirectoryRoleTemplate | Select-Object DisplayName, Description, Id | Sort-Object DisplayName

            # Format the result.
            $Result = foreach ($Policy in $CAPolicies) {
                $CustomObject = New-Object -TypeName psobject

                #Customer Name
                $CustomObject | Add-Member -MemberType NoteProperty -Name "CustomerName" -Value $CustomerName

                #current date
                $CustomObject | Add-Member -MemberType NoteProperty -Name "Date" -Value (Get-Date -Format 'yyyy-MM-dd')

                # displayName
                $CustomObject | Add-Member -MemberType NoteProperty -Name "displayName" -Value (Out-String -InputObject $Policy.displayName)


                # state
                $CustomObject | Add-Member -MemberType NoteProperty -Name "state" -Value (Out-String -InputObject $Policy.state)


                # includeUsers
                $Users = foreach ($User in $Policy.conditions.users.includeUsers) {
                    if ($User -ne 'All' -and $User -ne 'GuestsOrExternalUsers' -and $User -ne 'None') {
                        (Get-MgUser -Filter "id eq '$User'").userPrincipalName
                    }
                    else {
                        $User
                    }
                }

                $CustomObject | Add-Member -MemberType NoteProperty -Name "includeUsers" -Value (Out-String -InputObject $Users)


                # excludeUsers
                $Users = foreach ($User in $Policy.conditions.users.excludeUsers) {
                    if ($User -ne 'All' -and $User -ne 'GuestsOrExternalUsers' -and $User -ne 'None') {
                        (Get-MgUser -Filter "id eq '$User'").userPrincipalName
                    }
                    else {
                        $User
                    }
                }

                $CustomObject | Add-Member -MemberType NoteProperty -Name "excludeUsers" -Value (Out-String -InputObject $Users)


                # includeGroups
                $Groups = foreach ($Group in $Policy.conditions.users.includeGroups) {
                    if ($Group -ne 'All' -and $Group -ne 'None') {
                        (Get-MgGroup -Filter "id eq '$Group'").DisplayName
                    }
                    else {
                        $Group
                    }
                }

                $CustomObject | Add-Member -MemberType NoteProperty -Name "includeGroups" -Value (Out-String -InputObject $Groups)


                # excludeGroups
                $Groups = foreach ($Group in $Policy.conditions.users.excludeGroups) {
                    if ($Group -ne 'All' -and $Group -ne 'None') {
                        (Get-MgGroup -Filter "id eq '$Group'").DisplayName
                    }
                    else {
                        $Group
                    }
                }

                $CustomObject | Add-Member -MemberType NoteProperty -Name "excludeGroups" -Value (Out-String -InputObject $Groups)


                # includeRoles
                $Roles = foreach ($Role in $Policy.conditions.users.includeRoles) {
                    if ($Role -ne 'None' -and $Role -ne 'All') {
                        $RoleToCheck = ($EntraIDRoles | Where-Object { $_.Id -eq $Role }).displayName

                        if ($RoleToCheck) {
                            $RoleToCheck
                        }
                        else {
                            $Role
                        }
                    }
                    else {
                        $Role
                    }
                }

        $CustomObject | Add-Member -MemberType NoteProperty -Name "includeRoles" -Value (Out-String -InputObject $Roles)


        # excludeRoles
        $Roles = foreach ($Role in $Policy.conditions.users.excludeRoles) {
            if ($Role -ne 'None' -and $Role -ne 'All') {
                $RoleToCheck = ($EntraIDRoles | Where-Object { $_.Id -eq $Role }).displayName

                if ($RoleToCheck) {
                    $RoleToCheck
                }
                else {
                    $Role
                }
            }
            else {
                $Role
            }
        }

        $CustomObject | Add-Member -MemberType NoteProperty -Name "excludeRoles" -Value (Out-String -InputObject $Roles)


        # includeApplications
        $Applications = foreach ($Application in $Policy.conditions.applications.includeApplications) {
            if ($Application -ne 'None' -and $Application -ne 'All' -and $Application -ne 'Office365') {
                ($EnterpriseApps | Where-Object { $_.AppId -eq $Application }).displayName
            }
            else {
                $Application
            }
        }

        $CustomObject | Add-Member -MemberType NoteProperty -Name "includeApplications" -Value (Out-String -InputObject $Applications)


        # excludeApplications
        $Applications = foreach ($Application in $Policy.conditions.applications.excludeApplications) {
            if ($Application -ne 'None' -and $Application -ne 'All' -and $Application -ne 'Office365') {
                ($EnterpriseApps | Where-Object { $_.AppId -eq $Application }).displayName
            }
            else {
                $Application
            }
        }

        $CustomObject | Add-Member -MemberType NoteProperty -Name "excludeApplications" -Value (Out-String -InputObject $Applications)


        # includeUserActions
        $CustomObject | Add-Member -MemberType NoteProperty -Name "includeUserActions" -Value (Out-String -InputObject $Policy.conditions.applications.includeUserActions)


        # userRiskLevels
        $CustomObject | Add-Member -MemberType NoteProperty -Name "userRiskLevels" -Value (Out-String -InputObject $Policy.conditions.userRiskLevels)


        # signInRiskLevels
        $CustomObject | Add-Member -MemberType NoteProperty -Name "signInRiskLevels" -Value (Out-String -InputObject $Policy.conditions.signInRiskLevels)


        # includePlatforms
        $CustomObject | Add-Member -MemberType NoteProperty -Name "includePlatforms" -Value (Out-String -InputObject $Policy.conditions.platforms.includePlatforms)


        # excludePlatforms
        $CustomObject | Add-Member -MemberType NoteProperty -Name "excludePlatforms" -Value (Out-String -InputObject $Policy.conditions.platforms.excludePlatforms)


        # clientAppTypes
        $CustomObject | Add-Member -MemberType NoteProperty -Name "clientAppTypes" -Value (Out-String -InputObject $Policy.conditions.clientAppTypes)


        # includeLocations
        $includeLocations = foreach ($includeLocation in $Policy.conditions.locations.includeLocations) {
            if ($includeLocation -ne 'All' -and $includeLocation -ne 'AllTrusted' -and $includeLocation -ne '00000000-0000-0000-0000-000000000000') {
                (Get-MgIdentityConditionalAccessNamedLocation -Filter "Id eq '$includeLocation'").DisplayName
            }
            elseif ($includeLocation -eq '00000000-0000-0000-0000-000000000000') {
                'MFA Trusted IPs'
            }
            else {
                $includeLocation
            }
        }

        $CustomObject | Add-Member -MemberType NoteProperty -Name "includeLocations" -Value (Out-String -InputObject $includeLocations)


        # excludeLocation
        $excludeLocations = foreach ($excludeLocation in $Policy.conditions.locations.excludeLocations) {
            if ($excludeLocation -ne 'All' -and $excludeLocation -ne 'AllTrusted' -and $excludeLocation -ne '00000000-0000-0000-0000-000000000000') {
                (Get-MgIdentityConditionalAccessNamedLocation -Filter "Id eq '$includeLocation'").DisplayName
            }
            elseif ($excludeLocation -eq '00000000-0000-0000-0000-000000000000') {
                'MFA Trusted IPs'
            }
            else {
                $excludeLocation
            }
        }


        # excludeLocations
        $CustomObject | Add-Member -MemberType NoteProperty -Name "excludeLocations" -Value (Out-String -InputObject $excludeLocations)


        # grantControls
        $CustomObject | Add-Member -MemberType NoteProperty -Name "grantControls" -Value (Out-String -InputObject $Policy.grantControls.builtInControls)


        # termsOfUse
        $TermsOfUses = foreach ($TermsOfUse in $Policy.grantControls.termsOfUse) {
            $GraphUri = "https://graph.microsoft.com/v1.0/agreements/$TermsOfUse"
            (Get-MgAgreement | where Id -eq $TermsOfUse).displayName
        }
        
        $CustomObject | Add-Member -MemberType NoteProperty -Name "termsOfUse" -Value (Out-String -InputObject $TermsOfUses)


        # operator
        $CustomObject | Add-Member -MemberType NoteProperty -Name "operator" -Value (Out-String -InputObject $Policy.grantControls.operator)


        # sessionControlsapplicationEnforcedRestrictions
        $CustomObject | Add-Member -MemberType NoteProperty -Name "sessionControlsapplicationEnforcedRestrictions" -Value (Out-String -InputObject $Policy.sessionControls.applicationEnforcedRestrictions.isEnabled)


        # sessionControlscloudAppSecurity
        $CustomObject | Add-Member -MemberType NoteProperty -Name "sessionControlscloudAppSecurity" -Value (Out-String -InputObject $Policy.sessionControls.cloudAppSecurity.isEnabled)


        # sessionControlssignInFrequency
        $CustomObject | Add-Member -MemberType NoteProperty -Name "sessionControlssignInFrequency" -Value (Out-String -InputObject $Policy.sessionControls.signInFrequency)


        # sessionControlspersistentBrowser
        $CustomObject | Add-Member -MemberType NoteProperty -Name "sessionControlspersistentBrowser" -Value (Out-String -InputObject $Policy.sessionControls.persistentBrowser)


        # Return object.
        $CustomObject
    }

    # Add the result to the final result.
    $AllResults += $Result
        }
        catch{
            Write-Host "Failed to get Conditional Access Policies" -ForegroundColor Red
            #write the error code and message
            Write-Host $_.Exception.Message
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__ 
            $StatusCode = $_.Exception.Response.StatusCode.value__
            Write-Host $_
            #if 403 status code then display a message that this tenant does not appear to have required licensing
            if ($_.Exception.Response.StatusCode.value__ -eq 403){
                Write-Host "This tenant does not appear to have Entra ID P1 licensing" -ForegroundColor Yellow
            }
            Write-Log -customerName $CustomerName -EventName "Get-ConditionalAccessPolicies" -Status "Fail $StatusCode" -Message $_.Exception.Message
            return
        }

}

# Export the result to Excel.
Write-Verbose -Verbose -Message "Exporting report to Excel..."
$Path = "$((Get-Location).Path)\Conditional Access Policy Design Report $(Get-Date -Format 'yyyy-MM-dd').xlsx"
$AuditLogPath = "$((Get-Location).Path)\Conditional Access Policy Design Report $(Get-Date -Format 'yyyy-MM-dd') Audit Log.csv"
$AllResults | Export-Excel -Path $Path -WorksheetName "CA Policies" -BoldTopRow -FreezeTopRow -AutoFilter -AutoSize -ClearSheet -Show
$AuditLogs | Select-Object 'Customer Name', 'Timestamp', 'Event', 'Status', 'Message' | Export-CSV -Path $AuditLogPath -NoTypeInformation 

