#PARAMETERS
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



$ErrorActionPreference = "SilentlyContinue"

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

function Get-ExchangeToken ($CustomerTenantId) {
    $ExchangeTokenSplat = @{
        ApplicationId = $AppId # AppID in CSP tenant
        Scopes = 'https://outlook.office365.com/.default'
        ServicePrincipal = $true
        Credential = (New-Object System.Management.Automation.PSCredential ($AppId, (ConvertTo-SecureString $AppSecret -AsPlainText -Force)))
        RefreshToken = $RefreshToken
        Tenant = $CustomerTenantId # Customer TenantID
    }
    try{
        $ExchangeToken = New-PartnerAccessToken @ExchangeTokenSplat
        return $ExchangeToken.AccessToken
    }
    catch {
        $ErrorMsg = $_.Exception.Message
        Write-Host "Failed to get Exchange Token for $CustomerTenantId - $ErrorMsg"
        return
    }
}

# Initialize a list to hold the results
$Report = [System.Collections.Generic.List[Object]]::new()

# Function for adding output to the list
Function ExportCSV {
    Param($MailBoxName, $UPN, $InboxRule, $CustomerName)
    $Result = @{

            'Customer Name' = $CustomerName
            'Mailbox Name' = $MailBoxName
            'UPN' = $UPN
            'Inbox Rule Name' = $InboxRule.Name
            'Enabled' = $InboxRule.Enabled
            'Forward To' = $InboxRule.ForwardTo
            'Redirect To' = $InboxRule.RedirectTo
            'Forward As Attachment To' = $InboxRule.ForwardAsAttachmentTo
            'Move To Folder' = $InboxRule.MoveToFolder
            'Delete Message' = $InboxRule.DeleteMessage
            'Mark As Read' = $InboxRule.MarkAsRead
    
    }  
    $ReportLine = New-Object PSObject -Property $Result
    $Report.Add($ReportLine)
}

# Get the access token
$AccessToken = Get-GraphAccessToken -TenantId $TenantId
Write-Host "Got access token for Microsoft Graph" -ForegroundColor Green

# Define header with authorization token
$Headers = @{
    'Authorization' = "Bearer $AccessToken"
    'Content-Type'  = 'application/json'
}

# Get GDAP Customers
$GraphUrl = "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminCustomers"
$Customers = Invoke-RestMethod -Uri $GraphUrl  -Headers $Headers -Method Get
$Customers.value.displayName

#Loop Through Each Customer, Connect to Exchange Online and Get Mailboxes

$Customers.value | ForEach-Object {
    $Customer = $_
    $CustomerName = $Customer.displayName
    $CustomerTenantId = $Customer.id

    # Display the customer name
    Write-Host "Processing Customer: $CustomerName" -ForegroundColor Green

    # Get the access token for the customer
    try {
        Write-Host "Getting Exchange Token" -ForegroundColor Green
        $ExchangeAccessToken = Get-ExchangeToken -CustomerTenantId $CustomerTenantId
    }
    catch {
        $ErrorMSg = $_.Exception.Message
        Write-Host "Failed to get Exchange Token for $CustomerName - $ErrorMsg"
        return
    }

    try{
        #Connect to Exchange
    Connect-ExchangeOnline -DelegatedOrganization $CustomerTenantId -AccessToken $ExchangeAccessToken


    #Retrieve Inbox Rules for each mailbox
     Get-Mailbox -ResultSize unlimited | ForEach-Object {
        Write-Host "Processing Mailbox: $($_.DisplayName)"
        $MailBoxName = $_.DisplayName
        $UPN = $_.UserPrincipalName

        Get-InboxRule -MailBox $_.PrimarySmtpAddress | ForEach-Object {
            #if there are any inbox rules, export them to CSV and write host
            Write-Host "Exporting Inbox Rule: $($_.Name) for $UPN" -ForegroundColor Green
            ExportCSV -MailBoxName $MailBoxName -UPN $UPN -InboxRule $_ -CustomerName $CustomerName
        }
    }

    # Disconnect from the customer's Exchange Online session
    Disconnect-ExchangeOnline -Confirm:$false
    } catch {
        Write-Host "Failed to process $CustomerName - $_" -ForegroundColor Red
        Write-Host $_.Exception.Message
        return 
    }
}

# When displaying and exporting, select the properties in the desired order

$Report | Select-Object 'Customer Name', 'Mailbox Name', 'UPN', 'Inbox Rule Name', 'Enabled', 'Forward To', 'Redirect To', 'Forward As Attachment To', 'Move To Folder', 'Delete Message', 'Mark As Read' | Out-GridView

$OutputCsv = "$((Get-Location).Path)\Inbox Rules $(Get-Date -Format 'yyyy-MM-dd').csv"

$Report | Select-Object 'Customer Name', 'Mailbox Name', 'UPN', 'Inbox Rule Name', 'Enabled', 'Forward To', 'Redirect To', 'Forward As Attachment To', 'Move To Folder', 'Delete Message', 'Mark As Read' | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

Write-Host "Report saved to $OutputCsv" -ForegroundColor Cyan