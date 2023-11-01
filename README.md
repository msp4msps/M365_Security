# M365_Security
Repository for hosting security related PowerShell scripts.

[![github-follow](https://img.shields.io/github/followers/msp4msps?label=Follow&logoColor=purple&style=social)](https://github.com/msp4msps)
[![project-languages-used](https://img.shields.io/github/languages/count/msp4msps/tech_blog?color=important)](https://github.com/msp4msps/tech_blog)
[![project-top-language](https://img.shields.io/github/languages/top/msp4msps/tech_blog?color=blueviolet)](https://github.com/msp4msps/tech_blog)
[![license](https://img.shields.io/badge/License-MIT-brightgreen.svg)](https://choosealicense.com/licenses/mit/)

## Authenticating

Leverage the GetAccessToken.ps1 script to acquire an access token.

Follow my guide if you are unfamiliar with the Secure Application Model Process: https://tminus365.com/how-to-leverage-microsoft-apis-for-automation/


## API Permissions (Delegated)

Get-DisabledUsers
-AuditLog.Read.All,Directory.Read.All

Get-IntuneDevices
-DeviceManagementServiceConfig.Read.All, DeviceManagementManagedDevices.Read.All

## Instructions 

Get-IntuneDevices
-Run the script 
-Provide your Secure Application Model secrets to get an AccessToken
-Provide your Desired File Path for Output
-The Script will provide a CSV of all Intune Devices

<kbd>![screenshot1](Screenshots/IntuneDevices.jpg)</kbd>