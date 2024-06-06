#PARAMETERS
param ( 
    [Parameter(Mandatory=$false)]
    [string] $AppId, 
    [Parameter(Mandatory=$false)]
    [string] $AppSecret, 
    [Parameter(Mandatory=$false)]
    [string] $TenantId,
    [Parameter(Mandatory=$false)]
    [string] $redirectURI
)

# Step 1: Open the browser for user to authenticate and grant permissions
$authUrl = "https://login.microsoftonline.com/$tenantID/oauth2/v2.0/authorize?client_id=$AppId&response_type=code&redirect_uri=$redirectUri&response_mode=query&scope=https://graph.microsoft.com/.default"
Start-Process $authUrl

# Manual step: After the user authenticates, they will be redirected to the provided redirectUri with the authorization code in the query string. 
# Capture this code manually.
$authorizationCode = Read-Host "Please enter the authorization code from the URL"

# Step 3: Exchange the authorization code for tokens
$body = "grant_type=authorization_code&client_id=$AppID&client_secret=$AppSecret&code=$authorizationCode&redirect_uri=$redirectUri&scope=$scope"
$headers = @{ 'Content-Type' = 'application/x-www-form-urlencoded' }
$tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/token"
$response = Invoke-RestMethod -Method POST -Uri $tokenEndpoint -Body $body -Headers $headers


$accessToken = $response.access_token
$refreshToken = $response.refresh_token

# Output the tokens
Write-Host -ForegroundColor Green "Access token is" $accessToken
Write-Host -ForegroundColor Blue "Refresh token is"$refreshToken
