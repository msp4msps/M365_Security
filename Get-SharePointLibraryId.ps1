#PARAMETERS
param ( 
    [Parameter(Mandatory=$false)]
    [string] $tenant, 
    [Parameter(Mandatory=$false)]
    [string] $tenantId,
    [Parameter(Mandatory=$false)]
    [string] $siteName,
    [Parameter(Mandatory=$false)]
    [string] $docLib

)


#check for PnP PowerShell Module and install if not present
$PnPSite = Get-Module -Name PnP.PowerShell -ListAvailable
if ($null -eq $PnPSite) {
    Write-Host "PnP.PowerShell module not found, installing now"
    Install-Module -Name PnP.PowerShell -Force -AllowClobber
}
else {
    Write-Host "PnP.PowerShell module found"
    #Import the module
    Import-Module -Name PnP.PowerShell
}

# Example https://$tenant.sharepoint.com/sites/$siteName
$tenant = $tenant # tminus365com
$tenantId = $tenantId
$siteName = $siteName #Sharepoint site name
$docLib = $docLib #Sharepoint Document Library

#site URL
$siteUrl = "https://$tenant.sharepoint.com/sites/$siteName"

# Connection
try{
    Connect-PnPOnline -Url $siteUrl -Interactive
    Write-Host "Connected to SharePoint" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to SharePoint" -ForegroundColor Red
    Write-Host $_.Exception.Message
    break

}

# Convert Tenant ID
$tenantId = $tenantId -replace '-','%2D'

# Convert Site ID
$PnPSite = Get-PnPSite -Includes Id | select id
$PnPSite = $PnPSite.Id -replace '-','%2D'
$PnPSite = '%7B' + $PnPSite + '%7D'

# Convert Web ID
$PnPWeb = Get-PnPWeb -Includes Id | select id
$PnPWeb = $PnPWeb.Id -replace '-','%2D'
$PnPWeb = '%7B' + $PnPWeb + '%7D'

# Convert List ID
$PnPList = Get-PnPList $docLib -Includes Id | select id
$PnPList = $PnPList.Id -replace '-','%2D'

# Enumerate the Full URL
$FULLURL = 'tenantId=' + $tenantId + '&siteId=' + $PnPSite + '&webId=' + $PnPWeb + '&listId=' + $PnPList + '&webUrl=https%3A%2F%2F' + $tenant + '%2Esharepoint%2Ecom%2Fsites%2F' + $siteName + '&version=1'

# Output the FULL URL To Copy and Paste
Write-Output 'List ID: ' $FULLURL