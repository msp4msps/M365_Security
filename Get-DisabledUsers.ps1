<#
  .SYNOPSIS
  This script is used to garner inactive users in a tenant and output that information to a CSV file in the temp folder

  #>

  param (

  [Parameter(Mandatory=$true)]
  [string]$AccessToken # Access Token to call Microsoft Graph API
)



# Base URL for Microsoft Graph API

$graphApiUrl = "https://graph.microsoft.com/v1.0"



# Define Headers

$headers = @{

Authorization = "Bearer $AccessToken"

'Accept'      = 'application/json'

}





# Function to calculate if the last sign-in time is greater than 30 days

function IsLastSignInGreaterThan30Days($lastSignIn) {
$currentDate = Get-Date
try {
  $lastSignInDate = [DateTime]::Parse($lastSignIn)
  $daysDifference = ($currentDate - $lastSignInDate).Days
  return $daysDifference -gt 30
} catch {
  Write-Host "Failed to parse '$lastSignIn' as a DateTime."
  return $false
}
}



# Get users from Microsoft Graph sPI along with their sign-in activity

$usersEndpoint = "$graphApiUrl/users?`$select=displayName,userPrincipalName,signInActivity,accountEnabled,department,officeLocation&`$top=999"

$usersResponse = Invoke-RestMethod -Uri $usersEndpoint -Headers $headers -Method Get



# Initialize an array to store users who haven't signed in for over 30 days and are disabled

$inactiveDisabledUsers = @()



# Loop through each user and check if they haven't signed in for over 30 days

foreach ($user in $usersResponse.value) {

$signInActivities = $user.signInActivity

if ($signInActivities -ne $null) {

  $lastSignInTime = $signInActivities.lastSignInDateTime

  $isLastSignInGreaterThan30Days = IsLastSignInGreaterThan30Days $lastSignInTime


  # Calculate the number of days since the last sign-in activity

  $daysSinceLastSignIn = (Get-Date) - [DateTime]::Parse($lastSignInTime)


  if ($user.accountEnabled -and $isLastSignInGreaterThan30Days) {

      $AccountEnabledButNotActive = "True"

  } else {

      $AccountEnabledButNotActive = "False"

  }


  # Add the user to the array with the required information

  $inactiveDisabledUser = [PSCustomObject]@{

      DisplayName = $user.displayName

      UserPrincipalName = $user.userPrincipalName

      LastSignInTime = $lastSignInTime

      AccountEnabled = $user.accountEnabled

      SignInGreater30days = $isLastSignInGreaterThan30Days

      EnabledbutNotActive = $AccountEnabledButNotActive

      Department = $user.department

      OfficeLocation = $user.officeLocation

      DaysSinceLastSignIn = $daysSinceLastSignIn.Days

  }


  $inactiveDisabledUsers += $inactiveDisabledUser

}

}



# Export the results to a CSV file

$outputFilePath = "C:\Temp\DisabledUsers.csv"

$inactiveDisabledUsers | Export-Csv -Path $outputFilePath -NoTypeInformation



Write-Host "CSV file 'DisabledUsers.csv' has been generated."