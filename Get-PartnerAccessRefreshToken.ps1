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

# Construct authorization endpoint URL
$authEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize?client_id=$appId&response_type=code&redirect_uri=$redirectUri&scope=https://api.partnercenter.microsoft.com/.default"

# Navigate to authorization endpoint and obtain authorization code
Start-Process $authEndpoint
$code = Read-Host "Enter authorization code"


$body = "grant_type=authorization_code&client_id=$appId&client_secret=$appSecret&code=$code&redirect_uri=$redirectUri&scope=$scope"
$headers = @{ 'Content-Type' = 'application/x-www-form-urlencoded' }

$response = Invoke-RestMethod -Method POST -Uri $tokenEndpoint -Body $body -Headers $headers

return $response