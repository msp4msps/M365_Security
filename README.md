# M365_Security
Repository for hosting security related PowerShell scripts.

[![github-follow](https://img.shields.io/github/followers/msp4msps?label=Follow&logoColor=purple&style=social)](https://github.com/msp4msps)
[![project-languages-used](https://img.shields.io/github/languages/count/msp4msps/M365_Security?color=important)](https://github.com/msp4msps/M365_Security)
[![project-top-language](https://img.shields.io/github/languages/top/msp4msps/M365_Security?color=blueviolet)](https://github.com/msp4msps/M365_Security)
[![license](https://img.shields.io/badge/License-MIT-brightgreen.svg)](https://choosealicense.com/licenses/mit/)

## Authenticating

Leverage the GetAccessToken.ps1 script to acquire an access token.

Follow my guide if you are unfamiliar with the Secure Application Model Process: https://tminus365.com/how-to-leverage-microsoft-apis-for-automation/


## API Permissions (Delegated)

Get-DisabledUsers
- AuditLog.Read.All,Directory.Read.All

Get-IntuneDevices
- DeviceManagementServiceConfig.Read.All, DeviceManagementManagedDevices.Read.All

Get-DataProtectionSettings
-SharePointTenantSettings.Read.All

Get-CAPExcludedUsers
-Directory.Read.All,Policy.Read.All,GroupMember.ReadWrite.All

Get-MFAStats
-Directory.Read.All,Policy.Read.All,AuditLog.Read.All

## Instructions 

**Get Anonymous Links**
- Traverse to the file path where the script is located
- Run the file and specify the SharePoint URL you want to run the assessment on. Specify the name of the tenant you will be running the assessment on as well
- Sign in and consent with a Global admin in that tenant. 
-You will be prompted to review the report if there are any records returned. 

<kbd>![screenshot1](Screenshots/anonymous.jpg)</kbd>

**BECInvestigation**
- Traverse to the file path where the script is located
- Provide the user you want to investigate by UPN and enter the secrets + Tenant ID
- The Script will provide any log records related to recent MFA changes/registration, registered devices, or applications the user has consented to. 

<kbd>![screenshot1](Screenshots/BEC.jpg)</kbd>

**Get-IntuneDevices**
- Run the script 
- Provide your Secure Application Model secrets to get an AccessToken
- Provide your Desired File Path for Output
- The Script will provide a CSV of all Intune Devices

<kbd>![screenshot1](Screenshots/IntuneDevices.jpg)</kbd>

**Get-DataProtectionSettings**
-This Script collects SharePoint Settings, DLP Policies, Retention Policies, Information Protection Labels, and Label Policies
-Review the prerequisite of setting up the Secure Application Model prior to running the script to get the secrets you will need to acquire an access token. 
- Run the script 
- Provide a UPN of a Global Admin that can connect to Security and Compliance PowerShell ex: admin@novacoastschool.com (no quotes)
- Provide your Secure Application Model secrets to get an AccessToken
- The Script will provide a JSON of all policy information in the Temp folder of you C Drive


**Get-CAPExcludedUsers**
-This script records all users being excluded from each conditional access policy in a tenant. This includes direct exclusions or exclusions as part of a group. This script should help you identify any misconfigurations in policies. Generally, you should not have licensed users excluded in your policy definition. 
- CD into the file path you want to have the CSV file output
- Use the GetAccessToken.ps1 script to get an AccessToken for the script. 
- Provide your Secure Application Model secrets to get an AccessToken
- Run the script 
- The Script will output a CSV file in the location you navigate to


**Get-MFAStats**
-This script records all users MFA registration details including what primary MFA method a user leverages.Make sure you follow the readME if you have not already to get your Secure application model secrets.
- CD into the file path you want to have the CSV file output
- Run the script 
- Provide your Secure Application Model secrets to get an AccessToken
- The Script will output a CSV file in the location you navigate to