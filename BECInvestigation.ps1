#PARAMETERS
param ( 
    [Parameter(Mandatory=$true)]
    [string] $UserPrincipalName,
    [Parameter(Mandatory=$true)]
    [string] $tenantId,
    [Parameter(Mandatory=$true)]
    [string] $clientId,
    [Parameter(Mandatory=$true)]
    [string] $clientSecret
)

 <#
        .SYNOPSIS
            #Using the graph API we are going to retreieve the directory audit logs for a user specified in the input to help with Business email compromise investigation. The search will audit the users activity for up to 30 days.
        .DESCRIPTION
            Using the graph API we are going to retreieve the directory audit logs for a user specified in the input. We are going to look through the logs to find:
            - Changes to the Strong Authentication Methods (aka MFA Methods)
            - Changes to the Strong Authentication User Details (aka MFA Phone Numbers, Email, etc)
            - Consent to applications
            - Device Registration

        This report requires you have set up the necessary prerequisites for an application registration in the tenant or the Secure Applicable Model and GDAP in partner center.
        For more information on application registration: https://learn.microsoft.com/en-us/graph/auth-register-app-v2 
        For more information on GDAP automation, check out this blog: https://tminus365.com/my-automations-break-with-gdap-the-fix/
        API Docs: https://learn.microsoft.com/en-us/graph/api/directoryaudit-list?view=graph-rest-1.0&tabs=http 


        .PARAMETER
            ClientId
            The AppId of the application registration. AKA Client ID of the app registration.

        .PARAMETER
            clientSecret
            The AppSecret of the application registration. Aka the Client Secret on the app registration. 

        .PARAMETER 
            TenantId
            The tenant ID of the environment you want to run this against or your partner tenant ID if you are using the secure application model. 

        .PARAMETER 
            refreshToken
            The refreshToken of the Secure Applicable Model.
        
        .Minumum Requirements
            - PowerShell 5.1 or greater
            - Microsoft Graph Permissions on App Registration: AuditLog.Read.All
            - Min GDAP Permissions: Directory Reader

        .EXAMPLE
        #example call 
                #https://graph.microsoft.com/v1.0/auditLogs/directoryaudits?$filter=activityDisplayName eq 'Update user' AND initiatedBy/user/userPrincipalName eq 'msp4msps@tminus365.com'

                #example response 
                # {
                #     "id": "Directory_93a86bcd-c30a-4423-afa9-4c5f6a69f28d_XCJCZ_6018138",
                #     "category": "UserManagement",
                #     "correlationId": "93a86bcd-c30a-4423-afa9-4c5f6a69f28d",
                #     "result": "success",
                #     "resultReason": "",
                #     "activityDisplayName": "Update user",
                #     "activityDateTime": "2024-06-29T11:31:07.2852419Z",
                #     "loggedByService": "Core Directory",
                #     "operationType": "Update",
                #     "initiatedBy": {
                #         "app": null,
                #         "user": {
                #             "id": "d3a3bf9d-1caa-4044-8b94-1386f3df8038",
                #             "displayName": null,
                #             "userPrincipalName": "msp4msps@tminus365.com",
                #             "ipAddress": "",
                #             "userType": null,
                #             "homeTenantId": null,
                #             "homeTenantName": null
                #         }
                #     },
                #     "targetResources": [
                #         {
                #             "id": "d3a3bf9d-1caa-4044-8b94-1386f3df8038",
                #             "displayName": null,
                #             "type": "User",
                #             "userPrincipalName": "msp4msps@tminus365.com",
                #             "groupType": null,
                #             "modifiedProperties": [
                #                 {
                #                     "displayName": "StrongAuthenticationUserDetails",
                #                     "oldValue": "[{\"PhoneNumber\":null,\"AlternativePhoneNumber\":null,\"Email\":null,\"VoiceOnlyPhoneNumber\":null}]",
                #                     "newValue": "[{\"PhoneNumber\":\"+1 3523333960\",\"AlternativePhoneNumber\":null,\"Email\":null,\"VoiceOnlyPhoneNumber\":null}]"
                #                 },
                #                 {
                #                     "displayName": "Included Updated Properties",
                #                     "oldValue": null,
                #                     "newValue": "\"StrongAuthenticationUserDetails\""
                #                 },
                #                 {
                #                     "displayName": "TargetId.UserType",
                #                     "oldValue": null,
                #                     "newValue": "\"Member\""
                #                 }
                #             ]
                #         }
                #     ],
                #     "additionalDetails": [
                #         {
                #             "key": "UserType",
                #             "value": "Member"
                #         }
                #     ]
                # }


        .INPUTS
            Your ClientId, ClientSecret, Tenant ID, and UPN of the user you want to run it against. 

        .OUTPUTS
            Summary of the changes to the user's MFA methods and user details in the PowerShell window.

        .NOTES
            Author:   Nick Ross
            GitHub:   https://github.com/msp4msps/M365_Security
            Blog:     https://tminus365.com/
    
    #>

# Function to authenticate and get an access token
function Get-AccessToken {
    param (
        [string]$tenantId,
        [string]$clientId,
        [string]$clientSecret,
        [string]$scope
    )

    $authUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $body = @{
        grant_type    = "client_credentials"
        client_id     = $clientId
        client_secret = $clientSecret
        scope      = "https://graph.microsoft.com/.default"
    }

    $response = Invoke-RestMethod -Method POST -Uri $authUrl -Body $body -ContentType "application/x-www-form-urlencoded"

    return $response.access_token
}


# Define a list of authentication methods
$authenticationMethods = @(
    "TwoWayVoiceMobile", "TwoWaySms", "TwoWayVoiceOffice", "TwoWayVoiceOtherMobile",
    "TwoWaySmsOtherMobile", "OneWaySms", "PhoneAppNotification", "PhoneAppOTP"
)


# Get the access token
$accessToken = Get-AccessToken -tenantId $tenantId -clientId $clientId -clientSecret $clientSecret

# Define a function to replace strings
function Replace-Strings {
    param (
        [string]$string,
        [array]$oldValues,
        [array]$newValues
    )
    for ($i = 0; $i -lt $oldValues.Count; $i++) {
        $string = $string -replace [regex]::Escape($oldValues[$i]), $newValues[$i]
    }
    return $string
}

# Make REST API call to get directory audit logs
try{
    $url = "https://graph.microsoft.com/v1.0/auditLogs/directoryaudits?`$filter=activityDisplayName eq 'Update user' AND initiatedBy/user/userPrincipalName eq '$UserPrincipalName'"
    $auditLogs = Invoke-RestMethod -Uri $url -Headers @{Authorization= "Bearer $accessToken"} -Method Get    
} catch {
    Write-Host "Failed to get directory audit logs for $UserPrincipalName" -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit
}

#loop through the reponse and print out the modified properties for values where 
# Process the logs
$results = @()
foreach ($log in $auditLogs.value) {
    #Filter out the losgs to see just mofiied properties with StrongAuthenticationUserDetails
    # $modifiedProperties = $log.value.targetResources | Where-Object {$_.modifiedProperties -match "StrongAuthenticationUserDetails"} | Select-Object -ExpandProperty modifiedProperties

    $modifiedProperties = $log.targetResources.modifiedProperties
    if($modifiedProperties -eq $null) {
        Write-Host "No data found for User Update changes." -ForegroundColor Yellow
    } else {
        Write-Host "Data found for User Update changes" -ForegroundColor Green
    }

    foreach ($property in $modifiedProperties) {
        if($property.displayName -eq "StrongAuthenticationMethod") {
            Write-Host "Found a record in Strong Authentication Methods." -ForegroundColor Green
            $newValue = $property.NewValue | ConvertFrom-Json
            $oldValue = $property.OldValue | ConvertFrom-Json

            $oldMethods = $oldValue | Sort-Object MethodType
            $newMethods = $newValue | Sort-Object MethodType
            #Removed Value 
            foreach ($oldMethod in $oldMethods) {
                $matchedNewMethod = $newMethods | Where-Object {$_.MethodType -eq $oldMethod.MethodType}
                if ($null -eq $matchedNewMethod) {
                    $results += [PSCustomObject]@{
                        TimeGenerated = $log.ActivityDateTime
                        Action        = "Removed ($($authenticationMethods[$oldMethod.MethodType])) from Authentication Methods."
                        Actor         = $log.initiatedBy.user.userPrincipalName
                        Target        =  $log.targetResources.userPrincipalName
                        ChangedValue  = "Method Removed"
                        OldValue      = "$($oldMethod.MethodType): $($authenticationMethods[$oldMethod.MethodType])"
                        NewValue      = ""
                    }
                }
            }
              # Added Methods
              foreach ($newMethod in $newMethods) {
                $matchedOldMethod = $oldMethods | Where-Object {$_.MethodType -eq $newMethod.MethodType}
                if ($null -eq $matchedOldMethod) {
                    $results += [PSCustomObject]@{
                        TimeGenerated = $log.ActivityDateTime
                        Action        = "Added ($($authenticationMethods[$newMethod.MethodType])) as Authentication Method."
                        Actor         = $log.initiatedBy.user.userPrincipalName
                        Target        =  $log.targetResources.userPrincipalName
                        ChangedValue  = "Method Added"
                        OldValue      = ""
                        NewValue      = "$($newMethod.MethodType): $($authenticationMethods[$newMethod.MethodType])"
                    }
                }
            }
             # Default Method Changes
             $oldDefaultMethod = $oldMethods | Where-Object {$_.Default -eq $true}
             $newDefaultMethod = $newMethods | Where-Object {$_.Default -eq $true}
             
             if ($null -ne $oldDefaultMethod -and $null -ne $newDefaultMethod -and $oldDefaultMethod.MethodType -ne $newDefaultMethod.MethodType) {
                 $results += [PSCustomObject]@{
                     TimeGenerated = $log.ActivityDateTime
                     Action        = "Default Authentication Method was changed to ($($authenticationMethods[$newDefaultMethod.MethodType]))."
                     Actor         = $log.initiatedBy.user.userPrincipalName
                     Target        =  $log.targetResources.userPrincipalName
                     ChangedValue  = "Default Method"
                     OldValue      = "$($oldDefaultMethod.MethodType): $($authenticationMethods[$oldDefaultMethod.MethodType])"
                     NewValue      = "$($newDefaultMethod.MethodType): $($authenticationMethods[$newDefaultMethod.MethodType])"
                 }
             }
        }
        if ($property.displayName -eq "StrongAuthenticationUserDetails") {
            Write-Host "Found a record in Strong Authentication User Details." -ForegroundColor Green
            $newValue = $property.NewValue -replace "^\[|\]$"
            $oldValue = $property.OldValue -replace "^\[|\]$"
            $newValue = ConvertFrom-Json $newValue
            $oldValue = ConvertFrom-Json $oldValue
            #compare the old and new values, not just for phone numbers 
            $changedValues = @("PhoneNumber", "AlternativePhoneNumber", "Email", "VoiceOnlyPhoneNumber")
            foreach ($changedValue in $changedValues) {
                $oldValueStr = $oldValue.$changedValue
                $newValueStr = $newValue.$changedValue
                if ($oldValueStr -ne $newValueStr) {
                    write-host "Found a change in Strong Authentication." -ForegroundColor Green
                    $results += [PSCustomObject]@{
                        TimeGenerated    = $log.ActivityDateTime
                        Action       = "Changed $changedValue in Strong Authentication."
                        Actor        = $log.initiatedBy.user.userPrincipalName
                        Target       = $log.targetResources.userPrincipalName
                        ChangedValue = $changedValue
                        OldValue     = $oldValueStr
                        NewValue     = $newValueStr
                    }
                }
            }
        }
    }
}
try{
    $appConsentURL = "https://graph.microsoft.com/v1.0/auditLogs/directoryaudits?`$filter=activityDisplayName eq 'Consent to application' AND initiatedBy/user/userPrincipalName eq '$UserPrincipalName'"
    $consentLogs = Invoke-RestMethod -Uri $appConsentURL -Headers @{Authorization= "Bearer $accessToken"} -Method Get
} catch {
    Write-Host "Failed to get directory audit logs for $UserPrincipalName" -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit
}

#only run if there are logs
if($consentLogs -eq $null) {
    Write-Host "No data found for Consent to application" -ForegroundColor Yellow
} else {
    Write-Host "Data found for Consent to application" -ForegroundColor Green
    foreach ($log in $consentLogs.value) {
        $results += [PSCustomObject]@{
            TimeGenerated = $log.ActivityDateTime
            Action        = "Consented to application"
            Actor         = $log.initiatedBy.user.userPrincipalName
            Target        = $log.targetResources.userPrincipalName
            ChangedValue  = "Consented to application"
            OldValue      = ""
            NewValue      = "App Name: $($log.targetResources.displayName) on IP Address: $($log.initiatedBy.user.ipAddress)"
        }
    
    }
}

try{
    $deviceRegisterurl = "https://graph.microsoft.com/v1.0/auditLogs/directoryaudits?`$filter=activityDisplayName eq 'Register device' AND initiatedBy/user/userPrincipalName eq '$UserPrincipalName'"
    $registerLogs = Invoke-RestMethod -Uri $deviceRegisterurl -Headers @{Authorization= "Bearer $accessToken"} -Method Get
} catch {
    Write-Host "Failed to get directory audit logs for $UserPrincipalName" -ForegroundColor Red
    Write-Host $_.Exception.Message
    exit
}

#only run if there are logs
if($registerLogs -eq $null) {
    Write-Host "No data found for Register device" -ForegroundColor Yellow
} else {
    Write-Host "Data found for Register device" -ForegroundColor Green
    foreach ($log in $registerLogs.value) {

        if($log -ne $null) {
            #loop through log.additionalDetails and append key value pairs to a variable 
            $additionalDetails = $log.additionalDetails
            $additionalDetailsStr = ""
            foreach ($detail in $additionalDetails) {
                $additionalDetailsStr += "$($detail.key): $($detail.value) "
            }
    
            $results += [PSCustomObject]@{
                TimeGenerated = $log.ActivityDateTime
                Action        = "Register Device"
                Actor         = $log.initiatedBy.user.userPrincipalName
                Target        = $log.targetResources.userPrincipalName
                ChangedValue  = "RegisterDevice"
                OldValue      = ""
                NewValue      = $additionalDetailsStr
            }
        } else {
            Write-Host "No data found"
        }
    }
    
}

# Output the results
$results | Format-Table -AutoSize



