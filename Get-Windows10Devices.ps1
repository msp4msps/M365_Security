# Install the Microsoft Graph module if not already installed
if (!(Get-Module -Name Microsoft.Graph -ListAvailable)) {
    Install-Module Microsoft.Graph -Force -Scope CurrentUser
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Device.Read.All"

# Verify connection
$context = Get-MgContext
if (-not $context) {
    Write-Host "Failed to connect to Microsoft Graph. Please ensure proper permissions are granted." -ForegroundColor Red
    exit
}

# Query devices from Entra
Write-Host "Fetching device inventory..." -ForegroundColor Cyan
$devices = Get-MgDevice -All

if ($devices.Count -eq 0) {
    Write-Host "No devices found in Entra." -ForegroundColor Yellow
    exit
}

# Filter for devices not on Windows 11
$nonWindows11Devices = $devices | Where-Object {
    $_.OperatingSystem -eq "Windows" -and
    $_.OperatingSystemVersion -ne $null -and
    $_.AccountEnabled -eq $true -and
    [version]$_.OperatingSystemVersion -lt [version]"10.0.22000"
}

# Output results
if ($nonWindows11Devices.Count -gt 0) {
    Write-Host "Found the following non-Windows 11 devices:" -ForegroundColor Green
    $nonWindows11Devices | Select-Object DisplayName,ApproximateLastSignInDateTime, OperatingSystemVersion, DeviceOwnership, IsManaged, RegisteredUsers,RegisteredOwners | Format-Table -AutoSize
} else {
    Write-Host "No Windows 10 devices found in your Entra inventory." -ForegroundColor Yellow
}

# Disconnect from Microsoft Graph
Disconnect-MgGraph