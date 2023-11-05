function Get-AccessToken {
    param (
        [Parameter(Mandatory=$true)]
        [string]$clientID,

        [Parameter(Mandatory=$true)]
        [string]$clientSecret,

        [Parameter(Mandatory=$true)]
        [string]$tenantID = "common", # Your tenantID

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

$accessToken = Get-AccessToken

# Prompt the user for the Security Group ID
$SecurityGroup = Read-Host -Prompt "Enter the Security Group ID where your service principal is located"

# Prompt the user for the Role ID they wish to add to assignments
$RoleID = Read-Host -Prompt "Enter the Role ID you want to add to assignments"

# Initialize the RelationshipID variable with a blank value
$RelationshipID = ""

# Initialize the GroupAssignments variable as an empty array
$GroupAssignments = @()


# Define your Graph API endpoint for GDAP relationships
$gdapApiUrl = "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships?$filter=status eq 'active'"

# Use the existing access token for authorization
$headers = @{
    Authorization = "Bearer $accessToken"
    "Content-Type" = "application/json"
}

# Function to get GDAP relationships and access assignments
Function Get-GDAPAssignments {
    param (
        [string]$apiUrl,
        [hashtable]$headers
    )

    Try {
        # Make the API call to get the GDAP assignments
        $response = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get
        # Check if there are more pages of data
        while ($response.'@odata.nextLink') {
            # If there are more pages, update the apiUrl and perform another request
            $apiUrl = $response.'@odata.nextLink'
            $nextPage = Invoke-RestMethod -Uri $apiUrl -Headers $headers -Method Get
            $response.value += $nextPage.value
            $response.'@odata.nextLink' = $nextPage.'@odata.nextLink'
        }
        # Return the list of assignments
        return $response.value
    } Catch {
        Write-Error "Error fetching GDAP assignments: $_"
    }
}

# Call the function to get GDAP assignments
$activeGDAPRelationships = Get-GDAPAssignments -apiUrl $gdapApiUrl -headers $headers

# Display the assignments (for verification)
$activeGDAPRelationships | Format-Table


# Assume $gdapRelationships is the object containing all the GDAP relationships obtained from a previous API call
foreach ($gdapRelationship in $activeGDAPRelationships) {
    # Set the RelationshipID to the current GDAP Relationship ID
    $RelationshipID = $gdapRelationship.id

     # Store the customer's display name
     $customerDisplayName = $gdapRelationship.customer.displayName

    # Form the URI for the API call
    $uri = "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$RelationshipID/accessAssignments"

    # Make the API call
    $response = Invoke-RestMethod -Headers @{Authorization = "Bearer $accessToken"} -Uri $uri -Method Get

    # Check if any access assignments match the Security Group ID
    foreach ($assignment in $response.value) {
        if ($assignment.accessContainer.accessContainerId -eq $SecurityGroup) {
             $roleDefinitionIds = @()
             # Create an array to hold all roleDefinitionIds including the new one
             $roleDefinitionIds += $assignment.accessDetails.unifiedRoles | ForEach-Object { $_.roleDefinitionId }
             # Add the new RoleID to the array of roleDefinitionIds
             $roleDefinitionIds += $RoleID
 
            # Create a new object with the assignment ID and Relationship ID
            $assignmentObject = [PSCustomObject]@{
                AssignmentID    = $assignment.id
                RelationshipID  = $RelationshipID
                etag           = $assignment.'@odata.etag'
                CustomerName   = $customerDisplayName
                RoleDefinitionIds = $roleDefinitionIds
            }
            # Add the object to the GroupAssignments array
            $GroupAssignments += $assignmentObject
        }
    }
}

# Now $GroupAssignments array contains the assignments that match the Security Group

# Loop through each assignment in the GroupAssignments array
foreach ($groupAssignment in $GroupAssignments) {
    try {
        $headers = @{
            Authorization = "Bearer $accessToken"
            'If-Match' = $groupAssignment.ETag
            'Content-Type' = 'application/json'
        }
        $updateBody = @{
            accessDetails = @{
                unifiedRoles = $groupAssignment.RoleDefinitionIds | ForEach-Object {
                    @{ roleDefinitionId = $_ }
                }
            }
        }  | ConvertTo-Json -Depth 5
    

        $uri = "https://graph.microsoft.com/v1.0/tenantRelationships/delegatedAdminRelationships/$($groupAssignment.RelationshipID)/accessAssignments/$($groupAssignment.AssignmentID)"
        
        # Execute the PATCH request
        $response = Invoke-RestMethod -Headers $headers -Method PATCH -Uri $uri -Body $updateBody
        Write-Host "Success for $($groupAssignment.CustomerName) $($groupAssignment.AssignmentID)" -ForegroundColor Green
    } catch {
        Write-Host "Failed for $($groupAssignment.CustomerName) $($groupAssignment.AssignmentID): $($_.Exception.Message) check to see if this role is available as part of the GDAP relationship" -ForegroundColor Red
    }
}