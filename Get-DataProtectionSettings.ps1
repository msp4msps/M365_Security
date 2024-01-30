#Check for ExchangeOnlineModule 

if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Exchange Online Management Module not installed. Installing now..."
    Install-Module ExchangeOnlineManagement -Force
    Write-Host "Exchange Online Management Module installed."
}


#Assign a UPN for S&C PowerShell connection
$UPN = Read-host -Prompt "Enter the UPN of a global admin to connect to Security and Compliance PowerShell"

#Connect to Security and Compliance PowerShell in a try/catch block
try {
    Connect-IPPSSession -UserPrincipalName $UPN
    Write-Host "Connection to Security and Compliance PowerShell successful"
}
catch {
    Write-Host "Connection to Security and Compliance PowerShell failed. Please try again."
    break
}

#create an empty array to store the output
$outputAllDataProtection=@{}

function Get-AccessToken {
    param (
        [Parameter(Mandatory=$true)]
        [string]$clientID,

        [Parameter(Mandatory=$true)]
        [string]$clientSecret,

        [Parameter(Mandatory=$true)]
        [string]$tenantID, # Your tenantID

        [Parameter(Mandatory=$true)]
        [string]$refreshToken, # Your refreshToken

        [string]$scope = "https://graph.microsoft.com/.default" # Default scope for Microsoft Graph
    )

    # Token endpoint
    $tokenUrl = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/token"

    # Prepare the request body
    $body = @{
        client_id     = $clientID
        scope         = $scope
        client_secret = $clientSecret
        grant_type    = "refresh_token"
        refresh_token = $refreshToken
    }

    # Request the token in a try/catch block
    try {
        $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $body
    }
    catch {
        Write-Host "Error getting access token. Please try again."
        break
    }
    # Return the access token
    return $response.access_token
}

#Get Access Token
$accessToken = Get-AccessToken 

#Define the Graph API URI to call for SharePoint Settings
$uri = "https://graph.microsoft.com/v1.0/admin/sharepoint/settings"
$headers = @{
    "Authorization" = "Bearer $accessToken"
}

#Call the Graph API to get all SharePoint settings and convert the response to a PowerShell object
$sharePointSettings = Invoke-RestMethod -Uri $uri -Headers $headers

#Create a switch statement to redefine value for sharingCapability based on the value returned from the Graph API
switch ($sharePointSettings.sharingCapability) {
    "ExistingExternalUserSharingOnly" {$sharePointSettings.sharingCapability = "Existing Guest (Only guests already in your organization's directory.)"}
    "ExternalUserSharingOnly" {$sharePointSettings.sharingCapability = "New And Existing Guest (Guests must sign in or provide a verification code.)"}
    "ExternalUserAndGuestSharing" {$sharePointSettings.sharingCapability = "Anyone"}
    "Disabled" {$sharePointSettings.sharingCapability = "Only people in your organization (No external sharing allowed"}
}

$outputAllDataProtection | add-member -name "SharePointSettings" -value $sharePointSettings -MemberType NoteProperty

#Get the DLP Policies in the tenant
$DLPpolicies = Get-DlpCompliancePolicy | Select-Object @{N='DisplayName';E={$_.DisplayName}}, "Mode", "Comment", "WhenCreated", "Workload"


#Loop through DLP policies to get the rules for each policy and add them to $outputAllDataProtection object
foreach ($policy in $DLPpolicies) {
    $rules = Get-DlpComplianceRule -Policy $policy.DisplayName
    $policy.WhenCreated = $policy.WhenCreated.ToString("yyyy-MM-dd")
    $policy | Add-Member -MemberType NoteProperty -Name "Rules" -Value $rules.Name
}

$outputAllDataProtection | add-member -name "DlpPolicies" -value $DLPpolicies -MemberType NoteProperty

#Get the retention policies in the tenant
$retentionPolicies = Get-RetentionCompliancePolicy | Select-Object @{N='DisplayName';E={$_.Name}}, "Mode", "Comment", "WhenCreated", "Workload"

#Loop through retention policies to get the rules for each policy and add them to $outputAllDataProtection object
foreach ($policy in $retentionPolicies) {
    $rules = Get-RetentionComplianceRule -Policy $policy.DisplayName | Select-Object "RetentionComplianceAction", "RetentionDuration", "ExpirationDateOption"
    #Convert the retention duration from Days to years
    $rules.RetentionDuration = $rules.RetentionDuration / 365
    $policy.WhenCreated = $policy.WhenCreated.ToString("yyyy-MM-dd")
    #create a switch statement to convert the value for ExpirationDateOption to a new value with a default value of whatever is in the ExpirationDateOption property
    switch ($rules.ExpirationDateOption) {
        "ModificationAgeInDays" {$rules.ExpirationDateOption = "Based on last modified date"}
        "CreationAgeInDays" {$rules.ExpirationDateOption = "Based on created date"}
        default {$rules.ExpirationDateOption = $rules.ExpirationDateOption}
    }
    $policy | Add-Member -MemberType NoteProperty -Name "ComplianceAction" -Value $rules.RetentionComplianceAction
    $policy | Add-Member -MemberType NoteProperty -Name "RetentionDuration" -Value $rules.RetentionDuration
    $policy | Add-Member -MemberType NoteProperty -Name "ExpirationDateOption" -Value $rules.ExpirationDateOption
}

$outputAllDataProtection | add-member -name "retentionPolicies" -value $retentionPolicies -MemberType NoteProperty

#Get the information protection labels in the tenant
$infoProtectionLabels = Get-Label | Select-Object @{N='DisplayName';E={$_.DisplayName}}, "Mode", "ContentType", "Tooltip", "Workload", "WhenCreated"

#loop through each label and update the WhenCreated property to be in the correct format
foreach ($label in $infoProtectionLabels) {
    $label.WhenCreated = $label.WhenCreated.ToString("yyyy-MM-dd")
}

#Add the information protection labels to the $outputAllDataProtection object
$outputAllDataProtection | add-member -name "infoProtectionLabels" -value $infoProtectionLabels -MemberType NoteProperty

#Get the information protection label policies in the tenant
$infoProtectionLabelPolicies = Get-LabelPolicy | Select-Object @{N='DisplayName';E={$_.DistinguishedName}}, "Mode", "Type", "Comment", "Labels", "Workload", "Settings"

#Reassign the Displayname by splitting the commma and taking the first value 
foreach ($policy in $infoProtectionLabelPolicies) {
    $policy.DisplayName = $policy.DisplayName.Split(",")[0]
    $policy.DisplayName = $policy.DisplayName.Split("=")[1]
}

#Add the information protection label policies to the $outputAllDataProtection object
$outputAllDataProtection | add-member -name "infoProtectionLabelPolicies" -value $infoProtectionLabelPolicies -MemberType NoteProperty

#output all policies to json file
$outputAllDataProtection | ConvertTo-Json -Depth 10 | Out-File -FilePath "C:\temp\outputAllDataProtection.json" -Force