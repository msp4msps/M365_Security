#PARAMETERS
param ( 
    [Parameter(Mandatory=$false)]
    [string] $AppId, 
    [Parameter(Mandatory=$false)]
    [string] $AppSecret, 
    [Parameter(Mandatory=$false)]
    [string] $TenantId,
    [Parameter(Mandatory=$false)]
    [string] $refreshToken,
    [Parameter(Mandatory=$false)]
    [string] $PartnerRefreshToken,
    [Parameter(Mandatory=$false)]
    [string] $AppDisplayName
)


$ErrorActionPreference = "SilentlyContinue"

#See if Az.Accounts and Az.KeyVault Powershell module is installed and if not install it 
Write-Host "Checking for Az module"
$AzModule = Get-Module -Name Az.Accounts -ListAvailable
if ($null -eq $AzModule) {
    Write-Host "Az module not found, installing now"
    Install-Module -Name Az.Accounts -Force -AllowClobber
}
else {
    Write-Host "Az module found"
}

$AzModule = Get-Module -Name Az.KeyVault -ListAvailable
if ($null -eq $AzModule) {
    Write-Host "Az.KeyVault module not found, installing now"
    Install-Module -Name Az.KeyVault -Force -AllowClobber
}
else {
    Write-Host "Az.KeyVault module found"
}

#Connect to Azure, sign in with Global Admin Credentials
Connect-AzAccount

#Set values for Key Vault Name and Resource Group and Location 
$KeyVaultName = Read-Host -Prompt "Enter the name you want the key vault to be called. Make this unique"
$ResourceGroupName = Read-Host -Prompt "Enter the name of the resource group you want the key vault to be in"
$Location = Read-Host -Prompt "Enter the location you want the key vault to be in, i.e. East US"

#Create the Key Vault

try {
    $KeyVault = New-AzKeyVault -VaultName $KeyVaultName -ResourceGroupName $ResourceGroupName -Location $Location
    Write-Host "Key Vault created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create Key Vault" -ForegroundColor Red
    Write-Host $_.Exception.Message
    break
}

#Create the secrets based on the params in the key vault

#AppID
try{
    $secretvalue = ConvertTo-SecureString $AppId -AsPlainText -Force
    $secret = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name "AppID" -SecretValue $secretvalue
    Write-Host "AppID secret created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create AppID secret" -ForegroundColor Red
    Write-Host $_.Exception.Message
    return
}

#AppSecret
try{
    $secretvalue = ConvertTo-SecureString $AppSecret -AsPlainText -Force
    $secret = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name "AppSecret" -SecretValue $secretvalue
    Write-Host "AppSecret secret created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create AppSecret secret" -ForegroundColor Red
    Write-Host $_.Exception.Message
    return
}

#TenantID
try{
    $secretvalue = ConvertTo-SecureString $TenantId -AsPlainText -Force
    $secret = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name "TenantID" -SecretValue $secretvalue
    Write-Host "TenantID secret created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create TenantID secret" -ForegroundColor Red
    Write-Host $_.Exception.Message
    return
}

#RefreshToken
try{
    $secretvalue = ConvertTo-SecureString $refreshToken -AsPlainText -Force
    $secret = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name "RefreshToken" -SecretValue $secretvalue
    Write-Host "RefreshToken secret created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create RefreshToken secret" -ForegroundColor Red
    Write-Host $_.Exception.Message
    return
}

#PartnerRefreshToken
try{
    $secretvalue = ConvertTo-SecureString $PartnerRefreshToken -AsPlainText -Force
    $secret = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name "PartnerRefreshToken" -SecretValue $secretvalue
    Write-Host "PartnerRefreshToken secret created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create PartnerRefreshToken secret" -ForegroundColor Red
    Write-Host $_.Exception.Message
    return
}

#AppDisplayName
try{
    $secretvalue = ConvertTo-SecureString $AppDisplayName -AsPlainText -Force
    $secret = Set-AzKeyVaultSecret -VaultName $KeyVaultName -Name "AppDisplayName" -SecretValue $secretvalue
    Write-Host "AppDisplayName secret created successfully" -ForegroundColor Green
}
catch {
    Write-Host "Failed to create AppDisplayName secret" -ForegroundColor Red
    Write-Host $_.Exception.Message
    return
}

