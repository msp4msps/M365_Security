# Set up variables for the Graph API call
$accessToken = Get-AccessToken 
#Set up a variable to get the desired file path for the output files from the user
$FilePath = Read-Host -Prompt "Enter the desired file path for the output files. No paranthese or quotes. Example: C:\Users\username\Desktop"
$uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$top=999"
$autopilotDeviesURI = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$top=999"
$headers = @{
    "Authorization" = "Bearer $accessToken"
}

# Call the Graph API to get all Intune devices and convert the response to a PowerShell object
$response = Invoke-RestMethod -Uri $uri -Headers $headers
$devices = $response.value

# Call the Graph API to get all AutoPilot devices and convert the response to a PowerShell object
$response = Invoke-RestMethod -Uri $autopilotDeviesURI -Headers $headers
$autopilotDevices = $response.value

#Add to the data schema of the Intune device objects to include AutoPilot information and CheckInOver30Days
$devices | Add-Member -MemberType NoteProperty -Name "isAutopilotEnrolled" -Value $null
$devices | Add-Member -MemberType NoteProperty -Name "CheckInOver30Days" -Value $null
$device | Add-Member -MemberType NoteProperty -Name "StorageGreaterThan90%" -Value $null


# Create new device objects that match the data in the AutoPilot device objects to the Intune device objects. Intune devices objects do not have Autopilot information so it will need to be appended to the existing Intune device object
foreach ($device in $devices) {
    $autopilotDevice = $autopilotDevices | Where-Object { $_.id -eq $device.id }
    if ($autopilotDevice) {
        $device.isAutopilotEnrolled = $true
    } else {
        $device.isAutopilotEnrolled = $false
    }
}

# Create a table to store the device information
$table = @()

# Loop through each device and add its information to the table
foreach ($device in $devices) {
    $storageUsed = [math]::Round(($device.totalStorageSpaceInBytes - $device.freeStorageSpaceInBytes) / 1GB, 2)
    $storagePercent = [math]::Round(($storageUsed * 1GB / $device.totalStorageSpaceInBytes) * 100, 2)    
    $row = New-Object -TypeName PSObject -Property @{
        "Device ID" = $device.id
        "Device Name" = $device.deviceName
        "User" = $device.userDisplayName
        "Compliant" = $device.ComplianceState
        "AutoPilot Device" = $device.isAutopilotEnrolled
        "Enrollment Type" = $device.deviceEnrollmentType
        "Device Type" = $device.managedDeviceOwnerType
        "OS" = $device.operatingSystem
        "OS Version" = $device.osVersion
        "Encryption Enabled" = $device.isEncrypted
        "Last Check-In" = $device.lastSyncDateTime
        "CheckInOver30Days" = $device.lastSyncDateTime -lt (Get-Date).AddDays(-30)
        "Storage Used (GB)" = "$storageUsed GB"
        "Storage Percentage" = "$storagePercent"
        "StorageGreaterThan90%" = $storagePercent -gt 90
    }
    $table += $row
}

# Output the table to a CSV file in the desired order
$table | Select-Object "Device ID", "Device Name", "User", "Compliant", "AutoPilot Device", "Enrollment Type", "Device Type", "OS", "OS Version","Encryption Enabled", "Last Check-In", "CheckInOver30Days", "Storage Used (GB)", "Storage Percentage", "StorageGreaterThan90%" | Export-Csv -Encoding utf8 -NoTypeInformation -Path "$($FilePath)\IntuneDevices.csv"
