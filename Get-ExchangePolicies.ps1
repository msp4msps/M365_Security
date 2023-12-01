#Check for ExchangeOnlineModule 

if (!(Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Exchange Online Management Module not installed. Installing now..."
    Install-Module ExchangeOnlineManagement -Force
    Write-Host "Exchange Online Management Module installed."
}


try { 
    Connect-ExchangeOnline
    Write-Host "connection successfull"
} 
catch{ 
    # create response body in JSON format 
    $body = $_.Exception.Message | ConvertTo-Json -Compress -Depth 10 
    Write-Host  $body
    $outputAllExchangeParameters | add-member -name "ConnectionSuccesfull" -value "false" -MemberType NoteProperty
    break
}

# Define output parameter
$outputAllExchangeParameters=@{}



#Get Exchange TransportRules
$transportRule = Get-TransportRule | Select-Object @{N='RuleName';E={$_.Identity}}, "State", "Description"
$outputAllExchangeParameters | add-member -name "transportRule" -value $transportrule -MemberType NoteProperty


#Get AntiSpam Policies, keep in mind that some policies are default and dont have specific Filter Rules, these get the value Unknown. 
$hostedContentFilterPolicy = Get-HostedContentFilterPolicy | select-object -Property @{N='PolicyName';E={$_.Name}},"IsDefault","Enabled", @{N='BulkCompliantLevel(BCL)';E={$_.BulkThreshold}}, "HighConfidenceSpamAction","SpamAction","HighConfidencePhishAction",@{N='PhishAction';E={$_.PhishSpamAction}}, @{N='BulkMessageAction';E={$_.BulkSpamAction}}, @{N='RetentionOfSpamInQuarantine';E={$_.QuarantineRetentionPeriod}}, @{N='SpamSafetyTips';E={$_.InlineSafetyTipsEnabled}},"ZapEnabled", "SpamZapEnabled", "PhishZapEnabled", "Assignments" 

foreach ($rule in $($hostedContentFilterPolicy)) {
    try { $ruleSpecific = Get-HostedContentFilterRule $rule.PolicyName -ErrorAction Stop| Select-Object "State","Description"
    } 
    catch{ $ruleSpecific = @{"State"="Unknown"; "Description"= "Unknown"}
    } 
    $rule.Enabled = $ruleSpecific.State 
    $rule.Assignments = $ruleSpecific.Description
}
$outputAllExchangeParameters | add-member -name "hostedContentFilterPolicy" -value $hostedContentFilterPolicy -MemberType NoteProperty


#Get AntiMalware Policies, keep in mind that some policies are default and dont have specific Filter Rules, these get the value Unknown. 

$malwareFilterPolicy = Get-MalwareFilterPolicy | select-object -Property @{N='PolicyName';E={$_.Name}},"IsDefault","Enabled", @{N='EnableCommonAttachmentsFilter';E={$_.EnableFileFilter}}, "FileTypeAction","ZapEnabled",@{N='NotifyAdminsForUndeliveredFromInternalSenders';E={$_.EnableInternalSenderAdminNotifications}}, @{N='NotifyAdminsForUndeliveredFromExternalSenders';E={$_.EnableExternalSenderAdminNotifications}}, @{N='QuarantinePolicy';E={$_.QuarantineTag}}, "Assignments" 

foreach ($rule in $($malwareFilterPolicy)) {
    try { $ruleSpecific = Get-MalwareFilterRule $rule.PolicyName -ErrorAction Stop| Select-Object "State","Description"
    } 
    catch{ $ruleSpecific = @{"State"="Unknown"; "Description"= "Unknown"}
    } 
    $rule.Enabled = $ruleSpecific.State 
    $rule.Assignments = $ruleSpecific.Description
}
$outputAllExchangeParameters | add-member -name "malwareFilterPolicy" -value $malwareFilterPolicy -MemberType NoteProperty


# get safelink policy
$safeLinksPolicy = Get-SafeLinksPolicy | select-object -Property @{N='PolicyName';E={$_.Name}},"Enabled", @{N='SafeLinksForEmail';E={$_.EnableSafeLinksForEmail}}, @{N='SafeLinksForTeams';E={$_.EnableSafeLinksForTeams}}, @{N='SafeLinksForOfficeApps';E={$_.EnableSafeLinksForOffice}}, "AllowClickThrough", "TrackClicks", "Priority", "Assignments" 

foreach ($rule in $($safeLinksPolicy)) {
    try { $ruleSpecific = Get-SafeLinksRule $rule.PolicyName -ErrorAction Stop| Select-Object "State","Priority","Description"
    } 
    catch{ $ruleSpecific = @{"State"="Unknown"; "Priority"="Unknown"; "Description"= "Unknown"}
    } 
    $rule.Enabled = $ruleSpecific.State 
    $rule.Assignments = $ruleSpecific.Description
    $rule.Priority = $ruleSpecific.Priority
}
$outputAllExchangeParameters | add-member -name "safeLinksPolicy" -value $safeLinksPolicy -MemberType NoteProperty


# get safeAttachment policy
$safeAttachmentPolicy = Get-SafeAttachmentPolicy | select-object -Property @{N='PolicyName';E={$_.Name}},"Enabled", "Action", "QuarantineTag", "Priority", "Assignments" 

foreach ($rule in $($safeAttachmentPolicy)) {
    try { $ruleSpecific = Get-SafeAttachmentRule $rule.PolicyName -ErrorAction Stop| Select-Object "State","Priority","Description"
    } 
    catch{ $ruleSpecific = @{"State"="Unknown"; "Priority"="Unknown"; "Description"= "Unknown"}
    } 
    $rule.Enabled = $ruleSpecific.State 
    $rule.Assignments = $ruleSpecific.Description
    $rule.Priority = $ruleSpecific.Priority
}

$outputAllExchangeParameters | add-member -name "safeAttachmentPolicy" -value $safeAttachmentPolicy -MemberType NoteProperty

# get domainkeys (domain and enabled)
$DKIMSigningConfig=Get-DkimSigningConfig
$outputAllExchangeParameters | add-member -name "DKIMSigningConfig" -value $DKIMSigningConfig -MemberType NoteProperty


# get connectionfilter policy
$hostedConnectionFilterPolicy = Get-HostedConnectionFilterPolicy | select-object -Property @{N='PolicyName';E={$_.Name}},"isdefault", "IPAllowList", "IPBlockList", "EnableSafeList"
$outputAllExchangeParameters | add-member -name "hostedConnectionFilterPolicy" -value $hostedConnectionFilterPolicy -MemberType NoteProperty

#get auditing settings:
$adminAuditLogConfig= Get-AdminAuditLogConfig | select-object -Property @{N='AuditingEnabled';E={$_.UnifiedAuditLogIngestionEnabled}} 
$organizationConfig= Get-OrganizationConfig | select-object -Property @{N='MailboxAuditingEnabled';E={$_.AuditDisabled}}

$outputAllExchangeParameters | add-member -name "adminAuditLogConfig" -value $adminAuditLogConfig -MemberType NoteProperty
$outputAllExchangeParameters | add-member -name "organizationConfig" -value $organizationConfig -MemberType NoteProperty

#get antiPhish policy settings:
$antiPhishPolicy=Get-AntiPhishPolicy | Select-Object -Property @{N='PolicyName';E={$_.Name}}, "isDefault", "ImpersonationProtectionState", "EnableTargetedUserProtection","EnableMailboxIntelligenceProtection","EnableTargetedDomainsProtection","EnableOrganizationDomainsProtection","EnableMailboxIntelligence","EnableFirstContactSafetyTips","EnableSimilarUsersSafetyTips","EnableSimilarDomainsSafetyTips","EnableUnusualCharactersSafetyTips","EnableSpoofIntelligence","EnableSuspiciousSafetyTip","PhishThresholdLevel","TargetedUsersToProtect","ExcludedDomains","ExcludedSenders"
$outputAllExchangeParameters | add-member -name "antiPhishPolicy" -value $antiPhishPolicy -MemberType NoteProperty


#output all policies to json file
$outputAllExchangeParameters | ConvertTo-Json -Depth 10 | Out-File -FilePath "C:\temp\outputAllExchangeParameters.json" -Force

