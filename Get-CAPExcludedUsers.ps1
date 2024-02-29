    #The following script logs all the excluded users against Conditional Access Policies in your tenant.
    # Requires You to run the GetAccessToken.ps1 script to get the access token 


    param (
        [Parameter(Mandatory=$true)]
        [string]$AccessToken,

        [Parameter(Mandatory=$true)]
        [string]$CompanyId
        #Add the name of the company hered
    )


    function Get-Error {
        param (
            [string]$Message
        )
        Write-Error $Message
    }

    # Function to get user names by IDs
    function Get-UserNames {
        param (
            [string[]]$UserIds,
            [string]$AccessToken
        )
        $userNames = @()
        foreach ($userId in $UserIds) {
            $uri = "https://graph.microsoft.com/v1.0/users/$userId`?$select=displayName"
            try {
                $response = Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Bearer $AccessToken"} -Method Get
                if ($response.displayName) {
                    $userNames += $response.displayName            
            }
        }
            catch {
                Get-Error -Message $_.Exception.Message
            }
        }
        return $userNames
    }

    # Function to get excluded user names
    function Get-ExcludedUserNames {
        param (
            [object]$Users,
            [string]$AccessToken
        )
        $userNames = New-Object System.Collections.Generic.HashSet[string]
        if ($Users.excludeUsers.Count -ne 0 -and $Users.excludeUsers -notcontains 'All' -and $Users.excludeUsers -notcontains 'None') {
            $res = Get-UserNames -UserIds $Users.excludeUsers -AccessToken $AccessToken
            foreach ($each in $res) {
                $userNames.Add($each) | Out-Null
            }
        }

        if ($Users.excludeGroups.Count -ne 0) {
            foreach ($group in $Users.excludeGroups) {
                if ($group -match '^[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-4[0-9a-fA-F]{3}\-[89ABab][0-9a-fA-F]{3}\-[0-9a-fA-F]{12}$') { # UUID check
                    try {
                        $uri = "https://graph.microsoft.com/v1.0/groups/$group/members/microsoft.graph.user`?$select=displayName"
                        $response = Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Bearer $AccessToken"; "ConsistencyLevel" = "eventual"} -Method Get
                        foreach ($each in $response.value) {
                            $userNames.Add($each.displayName) | Out-Null
                        }
                    }
                    catch {
                        Get-Error -Message $_.Exception.Message
                    }
                }
            }
        }

        return [System.Linq.Enumerable]::ToArray($userNames)
    }

    # Function to get user IDs from groups
    function Get-UserIdsFromGroup {
        param (
            [string[]]$Groups,
            [string]$AccessToken
        )
        $userIDs = @()
        foreach ($group in $Groups) {
            try {
                $uri = "https://graph.microsoft.com/v1.0/groups/$group/members`?$select=id"
                $response = Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Bearer $AccessToken"} -Method Get
                if ($response.value.Count -gt 0) {
                    $userIDs += $response.value.id
                }
            }
            catch {
                Get-Error -Message $_.Exception.Message
            }
        }
        return $userIDs
    }

    # Function to fetch conditional access policies and process users
    function Get-ConditionalAccessPolicies {
        param (
            [string]$AccessToken,
            [string]$CompanyId
        )

        $policies = @()
        $hasNextPage = $true
        $requestUrl = 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'

        while ($hasNextPage) {
            try {
                $response = Invoke-RestMethod -Uri $requestUrl -Headers @{Authorization = "Bearer $AccessToken"} -Method Get
                if ($response.value -and $response.value.Count -gt 0) {
                    $policies += $response.value
                }
                else {
                    $hasNextPage = $false
                }
                if ($response.'@odata.nextLink') {
                    $requestUrl = $response.'@odata.nextLink'
                }
                else {
                    $hasNextPage = $false
                }
            }
            catch {
                Get-Error -Message $_.Exception.Message
                break
            }
        }

        $results = @()
        foreach ($policy in $policies) {
            $excludedUserIds = @()
            $licensedUsers = @()
            if ($policy.conditions.users.excludeUsers) {
                $excludedUserIds += $policy.conditions.users.excludeUsers
            }
            if ($policy.conditions.users.excludeGroups) {
                foreach ($groupId in $policy.conditions.users.excludeGroups) {
                    $userIds = Get-UserIdsFromGroup -Groups @($groupId) -AccessToken $AccessToken
                    $excludedUserIds += $userIds
                }
            }

            #Make API Call to see if excluded users are licensed
            foreach ($userId in $excludedUserIds) {
                $uri = "https://graph.microsoft.com/v1.0/users/$userId/licenseDetails"
                try {
                    $response = Invoke-RestMethod -Uri $uri -Headers @{Authorization = "Bearer $AccessToken"} -Method Get
                    if ($response.value -and $response.value.Count -gt 0) {
                        $licensedUsers += $userId
                    }
                }
                catch {
                    Get-Error -Message $_.Exception.Message
                }
            }
            $excludedUserNames = Get-ExcludedUserNames -Users $policy.conditions.users -AccessToken $AccessToken
            $licensedDisplayNames = Get-UserNames -UserIds $licensedUsers -AccessToken $AccessToken

            $result = [PSCustomObject]@{
                PolicyName = $policy.displayName
                State = $policy.state
                ExcludedUserIds = $excludedUserIds -join ', '
                ExcludedUserNames = $excludedUserNames -join ', '
                LicensedUsers = $licensedDisplayNames -join ', '
            }

            $results += $result
        }

        # Export result to CSV
        $csvPath = "ConditionalAccessPolicies-$CompanyId.csv"
        $results | Export-Csv -Path $csvPath -NoTypeInformation

        Write-Output "Exported conditional access policies to '$csvPath'"

        return $results
    }

    Get-ConditionalAccessPolicies -AccessToken $AccessToken -CompanyId $CompanyId