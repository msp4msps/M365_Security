param ( 
    [Parameter(Mandatory=$false)]
    [string] $SharePointUrl,
    [Parameter(Mandatory=$false)]
    [string] $TenantName
)

  <#
        .SYNOPSIS
            Automatically generate an Excel report containing any files with anonymous links in a SharePoint Document Library.
        .DESCRIPTION
            This script will automatically generate an Excel report containing any files with anonymous links in a SharePoint Document Library.
            The script will prompt for the SharePoint URL and then generate the report.
            The report will be saved in the same directory as the script.
            The report will contain the following columns:
            - Site Name
            - Library
            - File Name
            - File URL
            - Access Type
            - File Type
            - Link Expired Date
            - Days Since Expired
            - Link Created Date
            - Last Modified On
            - Shared Link

        .PARAMETER
            SharePointUrl
            The URL of the SharePoint site containing the Document Library.

            TenantName
            The name of the tenant. This is used to generate the report file name.

        .Minumum Requirements
            - SharePoint Online Management Shell
            - SharePoint Online PnP PowerShell Module
            - Excel
            - Windows PowerShell
            - SharePoint Online Tenant Admin Rights

        .OUTPUTS
            Excel Report

        .NOTES
            Author:   Nick Ross
            GitHub:   https://github.com/msp4msps/M365_Security
            Blog:     https://tminus365.com/
    
    #>

$ErrorActionPreference = "SilentlyContinue"

# Check if SharePoint PnP Management Shell is installed

$Module = Get-InstalledModule -Name PnP.PowerShell -RequiredVersion 1.12.0 -ErrorAction SilentlyContinue
If($Module -eq $null){
    Write-Host PnP PowerShell Module is not available -ForegroundColor Yellow
    $Confirm = Read-Host Are you sure you want to install module? [Yy] Yes [Nn] No
    If($Confirm -match "[yY]") { 
        Write-Host "Installing PnP PowerShell module..."
        Install-Module PnP.PowerShell -RequiredVersion 1.12.0 -Force -AllowClobber -Scope CurrentUser
        Import-Module -Name Pnp.Powershell -RequiredVersion 1.12.0           
    } 
    Else{ 
       Write-Host PnP PowerShell module is required to connect SharePoint Online.Please install module using Install-Module PnP.PowerShell cmdlet. 
       Exit
    }
}
Write-Host `nConnecting to SharePoint Online... -ForegroundColor Green

# Connect to SharePoint Online
try{
    Connect-PnPOnline -Url $SharePointUrl -Interactive
}
catch{
    Write-Host "Failed to connect to SharePoint Online. Please check the URL and try again." -ForegroundColor Red
    Exit
}


#Set Global Variables
$Global:ItemCount = 0
If($TenantName -eq ""){
    $TenantName = Read-Host "Enter your tenant name (e.g., 'contoso')"
}
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$ReportOutput = "$((Get-Location).Path)\$($TenantName)_SPO_Shared_Links $timestamp.csv"
$Site = (Get-PnPWeb | Select Title).Title

#Function to get shared links
Function Get-SharedLinkInfo($ListItems) {
 $Ctx = Get-PnPContext
  ForEach ($Item in $ListItems) {
    Write-Progress -Activity ("Site Name: $Site") -Status ("Processing Item: "+ $Item.FieldValues.FileLeafRef)
        $HasUniquePermissions = Get-PnPProperty -ClientObject $Item -Property HasUniqueRoleAssignments

        If ($HasUniquePermissions) {        
            $SharingInfo = [Microsoft.SharePoint.Client.ObjectSharingInformation]::GetObjectSharingInformation($Ctx, $Item, $false, $false, $false, $true, $true, $true, $true)
            $Ctx.Load($SharingInfo)
            $Ctx.ExecuteQuery()

            ForEach ($ShareLink in $SharingInfo.SharingLinks) {     
                Write-Host $ShareLink.HasExternalGuestInvitees
                If ($ShareLink.Url -and $ShareLink.LinkKind -like "*Anonymous*") { 
                    $LinkStatus = $true                     
                    $LinkCreated = ([DateTime]$ShareLink.Created).tolocalTime()
                    $CurrentDateTime = Get-Date
                    If($ShareLink.Expiration -ne ""){
                        $Expiration = ([DateTime]$ShareLink.Expiration).tolocalTime()
                        If($Expiration -lt $CurrentDateTime){
                            $daysExpired = ($currentDateTime - $expiration).Days
                            $LinkStatus = $true
                        } 
                    }
                    If($LinkStatus){
                        If($ShareLink.IsEditLink)
                        {
                            $AccessType="Write"
                        }
                        ElseIf($shareLink.IsReviewLink)
                        {
                            $AccessType="Review"
                        }
                        Else
                        {
                            $AccessType="Read"
                        }
                        $Results = [PSCustomObject]@{
                            "Site Name"             = $Site
                            "Library"          = $List.Title
                            "File Name"             = $Item.FieldValues.FileLeafRef
                            "File URL"         = $Item.FieldValues.FileRef
                            "Access Type"      = $AccessType
                            "File Type"         = $Item.FieldValues.File_x0020_Type 
                            "Link Expired Date"  = $Expiration
                            "Days Since Expired"     = $daysExpired
                            "Link Created Date "    = $LinkCreated   
                            "Last Modified On "   = ([DateTime]$ShareLink.LastModified).tolocalTime()                          
                            "Shared Link"       = $ShareLink.Url  
                        }
                        $Results | Export-CSV  -path $ReportOutput -NoTypeInformation -Append  -Force
                        $Global:ItemCount++
                    }                    
                }
            }
        }
    }
}

Function Get-SharedLinks{
    $ExcludedLists = @("Form Templates","Style Library","Site Assets","Site Pages", "Preservation Hold Library", "Pages", "Images",
                       "Site Collection Documents", "Site Collection Images")
    $DocumentLibraries = Get-PnPList | Where-Object {$_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.BaseType -eq "DocumentLibrary"}
    Foreach($List in $DocumentLibraries){
        $ListItems = Get-PnPListItem -List $List -PageSize 2000  | Where {$_.FileSystemObjectType -eq "File"}
        Get-SharedLinkInfo $ListItems
    }
}


try{
    Get-SharedLinks
}
catch{
    Write-Host "Failed to get shared links. Please check the URL and try again." -ForegroundColor Red
    Exit
}

if((Test-Path -Path $ReportOutput) -eq "True") 
{
    Write-Host `nThe output file contains $Global:ItemCount files
    Write-Host `n The Output file availble in:  -NoNewline -ForegroundColor Yellow
    Write-Host $ReportOutput
    $Prompt = New-Object -ComObject wscript.shell   
    $UserInput = $Prompt.popup("Do you want to open output file?",`   
    0,"Open Output File",4)   
    If ($UserInput -eq 6)   
    {   
        Invoke-Item "$ReportOutput"   
    } 
}
else{
    Write-Host -f Yellow "No Records Found"
}