param ( 
    [Parameter(Mandatory=$false)]
    [string] $KeyVaultName
)

  <#
        .SYNOPSIS
            Automatically generate an Excel report containing your current password policies and password expirations across users, across your customers that have a GDAP relationship.
            This is a fork from a legacy AdminDroid script that has been updated to use the Secure App Model for multitenancy use.
        .DESCRIPTION
            Uses Microsoft Graph to fetch all users and their password policies and expiration dates. This report requires you have set up the necessary prerequisites for the Secure Applicable Model and GDAP in partner center.
            For more information, check out this blog: https://tminus365.com/my-automations-break-with-gdap-the-fix/

        .PARAMETER
            KeyVaultName
            The keyvault name where the AppId, AppSecret, TenantId, and refreshToken are stored.

        .Minumum Requirements
            - PowerShell 5.1    
            - Microsoft Graph Permissions on App Registration: Directory.Read.All
            - Min GDAP Permissions: Directory Reader
            -Tenants must have Entra ID P1 licesning or higher


        .OUTPUTS
            Excel report with all users across tenants.

        .NOTES
            Author:   Nick Ross
            GitHub:   https://github.com/msp4msps/M365_Security
            Blog:     https://tminus365.com/
    
    #>




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

# Initialize a list to hold audit logs
$AuditLogs = [System.Collections.Generic.List[Object]]::new()

# Initialize a list to hold the results
$Report = [System.Collections.Generic.List[Object]]::new()

Function ExportCSV {
    Param($CustomerName, $User, $PwdLastChange, $PwdSinceLastSet, $PwdExpiryDate, $PwdExpireIn, $PwdExpiresIn, $LicenseStatus, $AccountStatus, $LastSignInDate, $InactiveDays)

    $Result = @{

            'Customer Name' = $CustomerName
            'Display Name' = $User.DisplayName
            'User Principal Name' = $UPN
            'Pwd Last Change Date' = $PwdLastChange
            'Days since Pwd Last Set' = $PwdSinceLastSet
            'Pwd Expiry Date' = $PwdExpiryDate
            'Friendly Expiry Time' = $PwdExpireIn
            'Days since Expiry(-) / Days to Expiry(+)' = $PwdExpiresIn
            'License Status' = $LicenseStatus
            'Account Status' = $AccountStatus
            'Last Sign-in Date' = $LastSignInDate
            'Inactive Days' = $InactiveDays
    
    }  
    $ReportLine = New-Object PSObject -Property $Result
    $Report.Add($ReportLine)
}


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
    Write-Host "Got partner access token for Microsoft Graph" -ForegroundColor Green
    Write-Log -customerName "Partner Tenant" -EventName "Get-GraphAccessToken" -Status "Success" -Message "Got partner access token for Microsoft Graph"

}
catch {
    Write-Host "Failed to partner access token for Microsoft Graph" -ForegroundColor Red
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

$Customers.value | ForEach-Object {
    $Customer = $_
    $CustomerName = $Customer.displayName
    $CustomerTenantId = $Customer.id

    # Display the customer name
    Write-Host "Processing Customer: $CustomerName" -ForegroundColor Green

    # Get the access token for the customer with Refresh token
    try {
        Write-Host "Getting Graph Token" -ForegroundColor Green
        $GraphAccessToken = Get-GraphAccessToken -TenantId $CustomerTenantId
        Write-Log -customerName $CustomerName -EventName "Get Graph Token for customer" -Status "Success" -Message "Got access token for Microsoft Graph"
    }
    catch {
        $ErrorMSg = $_.Exception.Message
        Write-Host "Failed to get Graph Token for $CustomerName - $ErrorMsg"
        Write-Log -customerName $CustomerName -EventName "Get Graph Token for customer" -Status "Fail" -Message $_.Exception.Message
        return
    }

    Write-Host "Connecting to MS Graph PowerShell..."

    try{
        #turn accesstoken to secure string
        $SecureString = ConvertTo-SecureString -String $GraphAccessToken -AsPlainText -Force
        Connect-MgGraph -AccessToken $SecureString
        Write-Host "Connected to MS Graph PowerShell" -ForegroundColor Green
        Write-Log -customerName $CustomerName -EventName "Connect to MS Graph PowerShell" -Status "Success" -Message "Connected to MS Graph PowerShell"
    }
    catch{
        Write-Host "Failed to connect to MS Graph PowerShell" -ForegroundColor Red
        Write-Log -customerName $CustomerName -EventName "Connect to MS Graph PowerShell" -Status "Fail" -Message $_.Exception.Message
        return
    }

$UserCount = 0 
$PrintedUser = 0
$PwdPolicy=@{}


#Getting Password policy for the domain
$Domains = Get-MgBetaDomain   #-Status Verified
foreach($Domain in $Domains)
{ 
    #Check for federated domain
    if($Domain.AuthenticationType -eq "Federated")
    {
        $PwdValidity = 0
    }
    else
    {
        $PwdValidity = $Domain.PasswordValidityPeriodInDays
        if($PwdValidity -eq $null)
        {
            $PwdValidity = 90
        }
    }
    $PwdPolicy.Add($Domain.Id,$PwdValidity)
}
Write-Host "Generating M365 users' password expiry report..." -ForegroundColor Magenta
#Loop through each user 
Get-MgBetaUser -All -Property DisplayName,UserPrincipalName,LastPasswordChangeDateTime,PasswordPolicies,AssignedLicenses,AccountEnabled,SigninActivity | foreach{ 
    $UPN = $_.UserPrincipalName
    $DisplayName = $_.DisplayName
    [boolean]$Federated = $false
    $UserCount++
    Write-Progress -Activity "`n     Processed user count: $UserCount "`n"  Currently Processing: $DisplayName"
    #Remove external users
    if($UPN -like "*#EXT#*")
    {
        return
    }
    $PwdLastChange = $_.LastPasswordChangeDateTime
    $PwdPolicies = $_.PasswordPolicies
    $LicenseStatus = $_.AssignedLicenses
    $LastSignInDate=$_.SignInActivity.LastSignInDateTime
    #Calculate Inactive days
    if($LastSignInDate -eq $null)
    { 
     $LastSignInDate="Never Logged-in"
     $InactiveDays= "-"
    }
    else
    {
     $InactiveDays= (New-TimeSpan -Start $LastSignInDate).Days
    }
    
    if($LicenseStatus -ne $null)
    {
        $LicenseStatus = "Licensed"
    }
    else
    {
        $LicenseStatus = "Unlicensed"
    }
    if($_.AccountEnabled -eq $true)
    {
        $AccountStatus = "Enabled"
    }
    else
    {
        $AccountStatus = "Disabled"
    }
    #Finding password validity period for user
    $UserDomain= $UPN -Split "@" | Select-Object -Last 1 
    $PwdValidityPeriod=$PwdPolicy[$UserDomain]
    #Check for Pwd never expires set from pwd policy
    if([int]$PwdValidityPeriod -eq 2147483647)
    {
        $PwdNeverExpire = $true
        $PwdExpireIn = "Never Expires"
        $PwdExpiryDate = "-"
        $PwdExpiresIn = "-"
    }
    elseif($PwdValidityPeriod -eq 0) #Users from federated domain
    {
        $Federated = $true
        $PwdExpireIn = "Insufficient data in O365"
        $PwdExpiryDate = "-"
        $PwdExpiresIn = "-"
    }
    elseif($PwdPolicies -eq "none" -or $PwdPolicies -eq "DisableStrongPassword") #Check for Pwd never expires set from Set-MsolUser
    {
        $PwdExpiryDate = $PwdLastChange.AddDays($PwdValidityPeriod)
        $PwdExpiresIn = (New-TimeSpan -Start (Get-Date) -End $PwdExpiryDate).Days
        if($PwdExpiresIn -gt 0)
        {
            $PwdExpireIn = "Will expire in $PwdExpiresIn days"
        }
        elseif($PwdExpiresIn -lt 0)
        {
            #Write-host `n $PwdExpiresIn
            $PwdExpireIn = $PwdExpiresIn * (-1)
            #Write-Host ************$pwdexpiresin
            $PwdExpireIn = "Expired $PwdExpireIn days ago"
        }
        else
        {
            $PwdExpireIn = "Today"
        }
    }
    else
    {
        $PwdExpireIn = "Never Expires"
        $PwdExpiryDate = "-"
        $PwdExpiresIn = "-"
    }
    #Calculating Password since last set
    $PwdSinceLastSet = (New-TimeSpan -Start $PwdLastChange).Days
   
    $PrintedUser++ 
    ExportCSV -CustomerName $CustomerName -User $_ -PwdLastChange $PwdLastChange -PwdSinceLastSet $PwdSinceLastSet -PwdExpiryDate $PwdExpiryDate -PwdExpireIn $PwdExpireIn -PwdExpiresIn $PwdExpiresIn -LicenseStatus $LicenseStatus -AccountStatus $AccountStatus -LastSignInDate $LastSignInDate -InactiveDays $InactiveDays
}
if($UserCount -eq 0)
{
    Write-Host No records found
}

Write-Host "`nThe output file contains " -NoNewline
Write-Host  $PrintedUser users-$customerName. -ForegroundColor Green


Disconnect-MgGraph | Out-Null

}

#Export the report to CSV
#Output file declaration 
$Location=Get-Location
$ExportCSV = "$Location\PasswordExpiryReport_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm-ss` tt).ToString()).csv" 
$Report | Select-Object 'Customer Name', 'Display Name', 'User Principal Name', 'Pwd Last Change Date', 'Days since Pwd Last Set', 'Pwd Expiry Date', 'Friendly Expiry Time', 'Days since Expiry(-) / Days to Expiry(+)', 'License Status', 'Account Status', 'Last Sign-in Date', 'Inactive Days' | Export-CSV -Path $ExportCSV -NoTypeInformation
$AuditLogPath = "$((Get-Location).Path)\Password $(Get-Date -Format 'yyyy-MM-dd') Audit Log.csv"
$AuditLogs | Select-Object 'Customer Name', 'Timestamp', 'Event', 'Status', 'Message' | Export-CSV -Path $AuditLogPath -NoTypeInformation 
Write-Host "Report saved to $Path" -ForegroundColor Cyan


if((Test-Path -Path $ExportCSV) -eq "True") 
{
    Write-Host `n "The Output file availble in:" -NoNewline -ForegroundColor Yellow; Write-Host "$ExportCSV" `n 

    $Prompt = New-Object -ComObject wscript.shell   
    $UserInput = $Prompt.popup("Do you want to open output file?",` 0,"Open Output File",4)   
    if ($UserInput -eq 6)   
    {   
        Invoke-Item "$ExportCSV"   
    } 
}