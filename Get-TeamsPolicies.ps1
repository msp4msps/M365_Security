#Check for ExchangeOnlineModule 

if (!(Get-Module -ListAvailable -Name MicrosoftTeams)) {
    Write-Host "Microsoft Teams Module not installed. Installing now..."
    Install-Module MicrosoftTeams -Force
    Write-Host "Microsoft Teams Module installed."
}


try { 
    Connect-MicrosoftTeams
    Write-Host "connection successfull"
} 
catch{ 
    # create response body in JSON format 
    $body = $_.Exception.Message | ConvertTo-Json -Compress -Depth 10 
    Write-Host  $body
    $outputAllTeamsParameters | add-member -name "ConnectionSuccesfull" -value "false" -MemberType NoteProperty
    break
}

# Define output parameter
$outputAllTeamsParameters=@{}

#Get Teams Application Permissions
$teamsAppPermission = Get-CsTeamsAppPermissionPolicy | Select-Object @{N='PolicyName';E={$_.Identity}}, @{N='MicrosoftApps';E={$_.DefaultCatalogAppsType}}, @{N='ThirdPartyApps';E={$_.GlobalCatalogAppsType}}, @{N='CustomApps';E={$_.PrivateCatalogAppsType}}
foreach ($policy in $teamsAppPermission) {
    # Check the ThirdPartyApps and CustomApps values
    if ($policy.ThirdPartyApps -ne "AllowedAppList" -or $policy.CustomApps -ne "AllowedAppList") {
        Write-Host "App Permissions: ThirdPartyApps and Custom apps should be restricted. The following policy should be updated: $($policy.PolicyName)" -ForegroundColor Red
    }
}
$outputAllTeamsParameters | add-member -name "teamsAppPermissions" -value $teamsAppPermission -MemberType NoteProperty

#Get Teams External Access Policy
$teamsExternalAccess = Get-CsTenantFederationConfiguration | Select-Object @{N='PolicyName';E={$_.Identity}}, "AllowedDomains", "BlockedDomains", "AllowTeamsConsumer", "AllowTeamsConsumerInbound", "AllowPublicUsers"
if ($teamsExternalAccess.AllowedDomains -like "AllowAllKnownDomains") {
    Write-Host "External Access: The external policy $($teamsExternalAccess.PolicyName) allows all outside domains. This should be restricted to specific domains." -ForegroundColor Red
}
$outputAllTeamsParameters | add-member -name "teamsExternalAccess" -value $teamsExternalAccess -MemberType NoteProperty

#Get Teams File Sharing Settings
$teamsFileSharing = Get-CsTeamsClientConfiguration | Select-Object @{N='PolicyName';E={$_.Identity}}, "AllowDropBox", "AllowBox", "AllowGoogleDrive", "AllowShareFile", "AllowEgnyte"
# Check if any of the file sharing options are enabled
if ($teamsFileSharing.AllowDropBox -eq $true -or $teamsFileSharing.AllowBox -eq $true -or $teamsFileSharing.AllowGoogleDrive -eq $true -or $teamsFileSharing.AllowShareFile -eq $true -or $teamsFileSharing.AllowEgnyte -eq $true) {
    Write-Host "File Sharing: 3rd party file sharing locations should be disabled." -ForegroundColor Red
}
$outputAllTeamsParameters | add-member -name "teamsFileSharing" -value $teamsFileSharing -MemberType NoteProperty

#Get Teams Meeting Policies
$teamsMeetingPolicies = Get-CsTeamsMeetingPolicy | Select-Object @{N='PolicyName';E={$_.Identity}},"AllowExternalParticipantGiveRequestControl", "AllowAnonymousUsersToStartMeeting", "AllowAnonymousUsersToJoinMeeting", "NewMeetingRecordingExpirationDays"
foreach ($policy in $teamsMeetingPolicies) {
    if ($policy.PolicyName -eq "Global" -and ($policy.AllowExternalParticipantGiveRequestControl -eq $true -or $policy.AllowAnonymousUsersToStartMeeting -eq $true -or $policy.AllowAnonymousUsersToJoinMeeting -eq $true)) {
        Write-Host "Meeting Policy: Your global meeting policy Allows External Participants to request screen control, or Allows for Anonymous users to start or join meetings. This policy should be updated." -ForegroundColor Red
    }
}

$outputAllTeamsParameters | add-member -name "teamsMeetingPolicies" -value $teamsMeetingPolicies -MemberType NoteProperty


#output all policies to json file
$outputAllTeamsParameters | ConvertTo-Json -Depth 10 | Out-File -FilePath "C:\temp\outputAllTeamsParameters.json" -Force

