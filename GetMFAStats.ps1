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

    # Request the token
    $response = Invoke-RestMethod -Method Post -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $body

    # Return the access token
    return $response.access_token
}

#Define variables
$clinentID = Read-Host -Prompt "Enter the client ID"
$clientSecret = Read-Host -Prompt "Enter the client secret"
$tenantID = Read-Host -Prompt "Enter the tenant ID"
$refreshToken = Read-Host -Prompt "Enter the refresh token"

# Get the access token
try {
    $AccessToken = Get-AccessToken -clientID $clinentID -clientSecret $clientSecret -tenantID $tenantID -refreshToken $refreshToken
}
catch {
    Write-Host "Failed to get access token: $_" -ForegroundColor Red
    break # Exit the script
}

# Graph API URL for MFA users (Modify accordingly to target the correct endpoint for MFA information)
$GraphApiUrl = "https://graph.microsoft.com/beta/reports/authenticationMethods/userRegistrationDetails"
$headers = @{
    Authorization = "Bearer $AccessToken"
}

# Get users
$Users = Invoke-RestMethod -Headers $headers -Uri $GraphApiUrl -Method Get

# Loop through users to get detailed information
$Users.value | ForEach-Object {
    $userDetailsUrl = "https://graph.microsoft.com/v1.0/users/$($_.id)?`$select=id,userPrincipalName,displayName, accountEnabled,jobTitle,department,assignedLicenses"
    $userDetails = Invoke-RestMethod -Headers $headers -Uri $userDetailsUrl -Method Get

    # Filter out if user has at least one assinged license
    if ($userDetails.assignedLicenses.Count -gt 0) {
        $assignedLicenses = $true 
    } else {
        $assignedLicenses = $false
    }

    # Extract the required information
    [PSCustomObject]@{
        UserID         = $_.id
        DisplayName    = $userDetails.displayName
        Email          = $_.userPrincipalName
        UserType       = $_.userType
        AccountEnabled = $userDetails.accountEnabled
        JobTitle       = $userDetails.jobTitle
        Department     = $userDetails.department
        MFAStatus      = $_.isMfaRegistered
        DefaultMFAMethod = $_.defaultMfaMethod
        #Convert from Object to String
        methodsRegistered = $_.methodsRegistered -join ","
        isSystemPreferredAuthenticationMethodEnabled = $_.isSystemPreferredAuthenticationMethodEnabled
        systemPreferredAuthenticationMethods = $_.systemPreferredAuthenticationMethods -join ","
        userPreferredMethodForSecondaryAuthentication = $_.userPreferredMethodForSecondaryAuthentication
        lastUpdatedDateTime = $_.lastUpdatedDateTime
        AdminUser       = $_.isAdmin
        PasswordlessCapable = $_.isPasswordlessCapable
        AssignedLicenses = $assignedLicenses
    }
} | Export-Csv -Path "MFA_Users_Details.csv" -NoTypeInformation