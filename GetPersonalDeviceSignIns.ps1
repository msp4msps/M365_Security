  <#
        .SYNOPSIS
            Automatically generate an Excel report containing any successful sign-ins from unmanaged devices in the last 30 days.
        .DESCRIPTION
            This script will automatically generate an Excel report containing any successful sign-ins from unmanaged devices in the last 30 days.
            The script will prompt for the Microsoft 365 tenant name and then generate the report.
            The report will be saved in the same directory as the script.
            The report will contain the following columns:
            - User Display Name
            - Apps
            - Locations
            - Is User Signing In From Non-US Location

        .PARAMETER 
            None
        
        .Minumum Requirements
            - Microsoft Graph PowerShell Module
            - Windows PowerShell 7 or later
            - Microsoft 365 Tenant Admin Rights

        .OUTPUTS
            Excel Report

        .NOTES
    #>
$ErrorActionPreference = "SilentlyContinue"

# Check if Microsoft Graph PowerShell Module is installed
$Module = Get-InstalledModule -Name Microsoft.Graph -RequiredVersion 1.6.0 -ErrorAction SilentlyContinue
If($Module -eq $null){
    Write-Host Microsoft Graph PowerShell Module is not available -ForegroundColor Yellow
    $Confirm = Read-Host Are you sure you want to install module? [Yy] Yes [Nn] No
    If($Confirm -match "[yY]") { 
        Install-Module -Name Microsoft.Graph -RequiredVersion 1.6.0 -Force -AllowClobber
    }
    Else {
        Write-Host Script cannot continue without the required module -ForegroundColor Red
        Exit
    }
}

# Calculate date range: end date is today, start date is 30 days ago
$endDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
$startDate = (Get-Date).AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ss.fffZ")

# Authenticate to Microsoft Graph
Connect-MgGraph -Scopes "AuditLog.Read.All"

# Initialize an empty list to store data
$allSignIns = @()

# Fetch sign-ins using Get-MgAuditLogSignIn and filter for unmanaged devices
$signIns = Get-MgAuditLogSignIn -Filter "createdDateTime ge $startDate and createdDateTime lt $endDate and (deviceDetail/isManaged eq false or deviceDetail/deviceId eq '') and status/errorCode eq 0" -PageSize 100

foreach ($record in $signIns) {
    # Create a custom object for each sign-in
    $signIn = [PSCustomObject]@{
        Id              = $record.Id
        Date            = $record.CreatedDateTime.ToString("MM-dd-yyyy")
        UserDisplayName = $record.userDisplayName
        UserPrincipalName = $record.UserPrincipalName
        ClientAppUsed   = $record.ClientAppUsed
        AppDisplayName  = $record.appDisplayName
        ConditionalAccessApplied = $record.ConditionalAccessStatus
        Location        = "$($record.location.city), $($record.location.state), $($record.location.countryOrRegion)"
        OperatingSystem = $record.deviceDetail.operatingSystem
        Browser         = $record.deviceDetail.browser
        IsManaged       = $record.deviceDetail.isManaged
    }
    $allSignIns += $signIn
}

# Group by user and aggregate location and app display name
$normalizedData = $allSignIns | Group-Object -Property UserDisplayName | ForEach-Object {
    [PSCustomObject]@{
        UserDisplayName = $_.Name
        Apps            = ($_.Group | Select-Object -ExpandProperty AppDisplayName | Sort-Object -Unique) -join ", "
        Locations       = ($_.Group | Select-Object -ExpandProperty Location | Sort-Object -Unique) -join ", "
    }
}

Write-Host "Unique Users Signing In From Unmanaged Devices in the Last 30 Days: $($normalizedData.Count)" -ForegroundColor Red

foreach ($record in $normalizedData) {
    Write-Output "User: $($record.UserDisplayName)"
    Write-Output "Apps: $($record.Apps)"
    Write-Output "Locations: $($record.Locations)"
    #example locaiton: Denver, Colorado, US, Sanford, Florida, US
    #look through the location data to see if US is present
    $outsideUS = $record.Locations -split ", " | Where-Object {$_ -eq "US"}
    if($outsideUS -ne "US"){
        Write-Output "User is signing in from a non-US location"
    }
    Write-Output ""
}

# Export the data to CSV, use get-path

$csvPath = "C:/temp/UserSignIns_UnmanagedDevices.csv"
$allSignIns | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

Write-Output "Data exported to $csvPath"
