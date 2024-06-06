#PARAMETERS
param ( 
    [Parameter(Mandatory=$false)]
    [string] $KeyVaultName,
    [Parameter(Mandatory=$false)]
    [string] $Permission
)

$ErrorActionPreference = "SilentlyContinue"

#See if Az.Accounts and Az.KeyVault Powershell module is installed and if not install it 
Write-Host "Checking for Az module"
$AzModule = Get-Module -Name Az.Accounts -ListAvailable
if ($null -eq $AzModule) {
    Write-Host "Az module not found, installing now"
    Install-Module -Name Az.Accounts -Force -AllowClobber
}
else {
    Write-Host "Az module found"
}

$AzModule = Get-Module -Name Az.KeyVault -ListAvailable
if ($null -eq $AzModule) {
    Write-Host "Az.KeyVault module not found, installing now"
    Install-Module -Name Az.KeyVault -Force -AllowClobber
}
else {
    Write-Host "Az.KeyVault module found"
}

#Connect to Azure, sign in with Global Admin Credentials
try{
    Connect-AzAccount
    Write-Host "Connected to Azure" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Azure" -ForegroundColor Red
    Write-Host $_.Exception.Message
    break

}


try {
    $AppId = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name "AppID" -AsPlainText
    Write-Host "Got AppID from Key Vault" -ForegroundColor Green
}
catch {
    Write-Host "Failed to get AppID from Key Vault" -ForegroundColor Red
    Write-Host $_.Exception.Message
    break
}

try {
    $AppSecret = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name "AppSecret" -AsPlainText
    Write-Host "Got AppSecret from Key Vault" -ForegroundColor Green
}
catch {
    Write-Host "Failed to get AppSecret from Key Vault" -ForegroundColor Red
    Write-Host $_.Exception.Message
    break
}

try {
    $TenantId = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name "TenantId" -AsPlainText
    Write-Host "Got TenantId from Key Vault" -ForegroundColor Green
}
catch {
    Write-Host "Failed to get TenantId from Key Vault" -ForegroundColor Red
    Write-Host $_.Exception.Message
    break
}

try {
    $refreshToken = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name "RefreshToken" -AsPlainText
    Write-Host "Got RefreshToken from Key Vault" -ForegroundColor Green
}
catch {
    Write-Host "Failed to get RefreshToken from Key Vault" -ForegroundColor Red
    Write-Host $_.Exception.Message
    break
}

try {
    $PartnerRefreshToken= Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name "PartnerRefreshToken" -AsPlainText
    Write-Host "Got Partner RefreshToken from Key Vault" -ForegroundColor Green
}
catch {
    Write-Host "Failed to get Partner Refresh from Key Vault" -ForegroundColor Red
    Write-Host $_.Exception.Message
    break
}

try {
    $AppDisplayName = Get-AzKeyVaultSecret -VaultName $KeyVaultName -Name "AppDisplayName" -AsPlainText
    Write-Host "Got App DisplayName from Key Vault" -ForegroundColor Green
}
catch {
    Write-Host "Failed to get AppDisplay Name from Key Vault" -ForegroundColor Red
    Write-Host $_.Exception.Message
    break
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
    Write-Host "Got accestoken for $AppID"
    return $AuthResponse.access_token
}

function Get-PartnerAccessToken ($TenantId) {

    $Body = @{
        'tenant'        = $TenantId
        'client_id'     = $AppId
        'scope'         = 'https://api.partnercenter.microsoft.com/.default'
        'client_secret' = $AppSecret
        'grant_type'    = 'refresh_token'
        'refresh_token' = $PartnerRefreshToken
    }
    $Params = @{
        'Uri'         = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        'Method'      = 'Post'
        'ContentType' = 'application/x-www-form-urlencoded'
        'Body'        = $Body
    }
    $AuthResponse = Invoke-RestMethod @Params
    Write-Host "Got partner accestoken for $AppID"
    return $AuthResponse.access_token
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

# Get the access token
try {
    $AccessToken =  Get-GraphAccessToken -TenantId $TenantId
    Write-Host "Got  access token for Microsoft Graph" -ForegroundColor Green
    Write-Log -customerName "Partner Tenant" -EventName "Get-GraphAccessToken" -Status "Success" -Message "Got access token for Microsoft Graph"

}
catch {
    Write-Host "Failed to get access token for Microsoft Graph" -ForegroundColor Green
    $ErrorMessage = $_.Exception.Message
    Write-Log -customerName "Partner Tenant" -EventName "Get-GraphAccessToken" -Status "Fail" -Message $ErrorMessage
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
    Write-Log -customerName "Partner Tenant" -EventName "Get-GDAPCustomers" -Status "Success" -Message "Got GDAP Customers"

}catch{
    Write-Host "Failed to get GDAP Customers" -ForegroundColor Red
    $ErrorMessage = $_.Exception.Message
    Write-Log -customerName "Partner Tenant" -EventName "Get-GDAPCustomers" -Status "Fail" -Message $ErrorMessage
    break

}

# Get the partner access token
try {
    $PartnerAccessToken =  Get-PartnerAccessToken -TenantId $TenantId
    Write-Host "Got partner access token" -ForegroundColor Green
    Write-Log -customerName "Partner Tenant" -EventName "Get-PartnerAccessToken" -Status "Success" -Message "Got partner access token"

}
catch {
    Write-Host "Failed to retrieve partner access token" -ForegroundColor Red
    Write-host $ErrorMessage = $_.Exception.Message
    Write-Log -customerName "Partner Tenant" -EventName "Get-GraphAccessToken" -Status "Fail" -Message $ErrorMessage
    break
}

$Partnerheaders = @{
    Authorization = "Bearer $($PartnerAccessToken)"
    'Accept'      = 'application/json'
}

foreach ($Customer in $Customers.value) {
    $CustomerTenantId = $Customer.id
    $customerName = $Customer.displayName
    Write-Host "Processing $customerName"
    $removeUri = "https://api.partnercenter.microsoft.com/v1/customers/$CustomerTenantId/applicationconsents/$AppId"
    #Remove existing consent
    try{
        Invoke-RestMethod -Uri $removeUri -Headers $Partnerheaders -Method DELETE -ContentType 'application/json'
        Write-Host "Successfuly Removed Consent from $customerName"
    } catch {
        write-host "Failed to remove consent from $customerName -  $_.Exception.Message"
    }

    try{

        # Consent to required applications
        $uri = "https://api.partnercenter.microsoft.com/v1/customers/$CustomerTenantId/applicationconsents"
        $body = @{
            applicationGrants = @(
                @{
                    enterpriseApplicationId = "00000003-0000-0000-c000-000000000000"
                    scope                   = $Permission
                },
                @{
                    enterpriseApplicationId = "00000002-0000-0ff1-ce00-000000000000"
                    scope                   = "Exchange.Manage"
                }
            )
            applicationId   = $AppId
            displayName     = $AppDisplayName
        } | ConvertTo-Json
      Write-Host $body 
      Write-Host $Permission
      Invoke-RestMethod -Uri $uri -Headers $Partnerheaders -Method POST -Body $body -ContentType 'application/json'
    } catch {
        write-host "Failed to add consent from $customerName -  $_.Exception.Message"
    }

  }				