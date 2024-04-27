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

function Get-AppPermAccessToken ($TenantId) {
    $Body = @{
        'tenant'        = $TenantId
        'client_id'     = $AppId
        'scope'         = 'https://graph.microsoft.com/.default'
        'client_secret' = $AppSecret
        'grant_type'    = 'client_credentials'
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


# Initialize a list to hold the results
$Report = [System.Collections.Generic.List[Object]]::new()


# Function for adding output to the list
Function ExportCSV {
    Param($MailBoxName, $UPN, $InboxRule, $CustomerName)

    if ($InboxRule.actions.ForwardTo.Count -gt 0) {
        # Collect email addresses into an array
        $emailAddresses = $InboxRule.actions.ForwardTo | ForEach-Object { $_.emailaddress.address }
    
         # Join email addresses with a comma
         $ForwardTo = $emailAddresses -join ','
    }
    else {
        $ForwardTo = $null
    }
    if ($InboxRule.actions.redirectTo.Count -gt 0) {
        #loop through the fw to addresses and join them with a comma
        $redirects = $InboxRule.actions.RedirectTo | ForEach-Object { $_.emailaddress.address }
        $RedirectTo = $redirects -join ','
    }
    else {
        $RedirectTo = $null
    }
    if ($InboxRule.actions.forwardAsAttachmentTo.Count -gt 0) {
        #loop through the fw to addresses and join them with a comma
        $forwardAsAttachmentTo = $InboxRule.actions.forwardAsAttachmentTo | ForEach-Object { $_.emailaddress.address }
        $forwardAsAttachmentTo = $forwardAsAttachmentTo -join ','
    }
    else {
        $forwardAsAttachmentTo = $null
    }

    If ($InboxRule.actions.MoveToFolder -ne $null) {
        $MoveToFolder = $true
    }
    else {
        $MoveToFolder = $null
    }

    $Result = @{

            'Customer Name' = $CustomerName
            'Mailbox Name' = $MailBoxName
            'UPN' = $UPN
            'Inbox Rule Name' = $InboxRule.displayName
            'Enabled' = $InboxRule.isEnabled
            'Forward To' = $ForwardTo
            'Redirect To' = $RedirectTo
            'Forward As Attachment To' = $forwardAsAttachmentTo
            'Move To Folder' = $MoveToFolder
            'Delete Message' = $InboxRule.actions.delete
            'Mark As Read' = $InboxRule.actions.markAsRead
    
    }  
    $ReportLine = New-Object PSObject -Property $Result
    $Report.Add($ReportLine)
}

# Get the access token
try {
    $AccessToken =  Get-GraphAccessToken -TenantId $TenantId
    Write-Host "Got partner access token for Microsoft Graph" -ForegroundColor Green
    Write-Log -customerName "Partner Tenant" -EventName "Get-GraphAccessToken" -Status "Success" -Message "Got partner access token for Microsoft Graph"

}
catch {
    Write-Host "Got partner access token for Microsoft Graph" -ForegroundColor Green
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

#Loop Through Each Customer, Connect to Exchange Online and Get Mailboxes

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

    
    # Get the access token for the customer with client credentials
    try {
        Write-Host "Getting Application Permission Graph Token" -ForegroundColor Green
        $AppPermAccessToken = Get-AppPermAccessToken -TenantId $CustomerTenantId
        Write-Log -customerName $CustomerName -EventName "Get Application Permission Graph Token for customer" -Status "Success" -Message "Got access token for Microsoft Graph"
    }
    catch {
        $ErrorMSg = $_.Exception.Message
        Write-Host "Failed to get Application Permission Graph Token for $CustomerName - $ErrorMsg"
        Write-Log -customerName $CustomerName -EventName "Get Application Permission Graph Token for customer" -Status "Fail" -Message $_.Exception.Message
        return
    }


    try{
    # Use the graph access token to call the graph API to get all mailboxes

    $GraphUrl = "https://graph.microsoft.com/beta/reports/getMailboxUsageDetail(period='D7')?`$format=application/json"
    $Headers = @{
        'Authorization' = "Bearer $GraphAccessToken"
        'Content-Type'  = 'application/json'
    }
    $Mailboxes = Invoke-RestMethod -Uri $GraphUrl  -Headers $Headers -Method Get
    Write-Log -customerName $CustomerName -EventName "Get Mailboxes" -Status "Success" -Message "Retrieved mailboxes"

    } catch {
        $ErrorMSg = $_.Exception.Message
        Write-Host "Failed to get mailboxes for $CustomerName - $ErrorMsg" -ForegroundColor Red
        Write-Log -customerName $CustomerName -EventName "Get Mailboxes" -Status "Fail" -Message $_.Exception.Message
        return
    }
    #Reset headers for the application permission token
    $Headers = @{
        'Authorization' = "Bearer $AppPermAccessToken"
        'Content-Type'  = 'application/json'
    }
    #loop thright the mailboxes to get the inbox rules
    $Mailboxes.value | ForEach-Object {
        Write-Host "Processing Mailbox: $($_.displayName)"
        $MailBoxName = $_.DisplayName
        $UPN = $_.UserPrincipalName
        try{
            $GraphInboxRuleURL = "https://graph.microsoft.com/v1.0/users/$UPN/mailFolders/inbox/messageRules"
            $InboxRules = (Invoke-RestMethod -Uri $GraphInboxRuleURL -Headers $Headers -Method Get).value
            #if there are any inbox rules, loop through them and export them to CSV
            if ($InboxRules.Count -gt 0) {
                $InboxRules | ForEach-Object {
                    Write-Host "Exporting Inbox Rule: $($_.displayName) for $UPN" -ForegroundColor Green
                    ExportCSV -MailBoxName $MailBoxName -UPN $UPN -InboxRule $_ -CustomerName $CustomerName
                }
                Write-Log -customerName $CustomerName -EventName "Get Inbox Rules-$UPN" -Status "Success" -Message "Retrieved inbox rules"
            } else{
                Write-Log -customerName $CustomerName -EventName "Get Inbox Rules-$UPN" -Status "Success" -Message "No Inbox rules for $UPN"
            }
        } catch {
            $ErrorMSg = $_.Exception.Message
            Write-Host "Failed to get inbox rules for $UPN - $ErrorMsg" -ForegroundColor Red
            Write-Log -customerName $CustomerName -EventName "Get Inbox Rules-$UPN" -Status "Fail" -Message $_.Exception.Message
            return
        }
    }

} 

# When displaying and exporting, select the properties in the desired order

$Report | Select-Object 'Customer Name', 'Mailbox Name', 'UPN', 'Inbox Rule Name', 'Enabled', 'Forward To', 'Redirect To', 'Forward As Attachment To', 'Move To Folder', 'Delete Message', 'Mark As Read' | Out-GridView
$path = "$((Get-Location).Path)\Inbox Rules $(Get-Date -Format 'yyyy-MM-dd').csv"
$Report | Select-Object 'Customer Name', 'Mailbox Name', 'UPN', 'Inbox Rule Name', 'Enabled', 'Forward To', 'Redirect To', 'Forward As Attachment To', 'Move To Folder', 'Delete Message', 'Mark As Read' | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
$AuditLogPath = "$((Get-Location).Path)\Inbox Rules $(Get-Date -Format 'yyyy-MM-dd') Audit Log.csv"
$AuditLogs | Select-Object 'Customer Name', 'Timestamp', 'Event', 'Status', 'Message' | Export-CSV -Path $AuditLogPath -NoTypeInformation 
Write-Host "Report saved to $Path" -ForegroundColor Cyan