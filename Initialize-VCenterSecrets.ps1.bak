#Requires -Modules Microsoft.PowerShell.SecretManagement, Microsoft.PowerShell.SecretStore

<#
.SYNOPSIS
    One-time setup script to configure the SecretStore vault and store vCenter credentials.

.DESCRIPTION
    Reads the vCenter list from config/vcenters.json and prompts for credentials for each
    vCenter server. Credentials are stored encrypted in a local SecretStore vault using
    Microsoft.PowerShell.SecretManagement.

    This script must be run interactively (it prompts for credentials) and must be run
    as the same user account that will execute the inventory scripts (e.g., the Task
    Scheduler service account).

.PARAMETER ConfigFile
    Path to the JSON configuration file containing the vCenter list.
    Defaults to config/vcenters.json relative to this script's directory.

.EXAMPLE
    .\Initialize-VCenterSecrets.ps1
    Prompts for credentials for each vCenter defined in config/vcenters.json.

.EXAMPLE
    .\Initialize-VCenterSecrets.ps1 -ConfigFile "C:\Config\vcenters.json"
    Uses a custom config file path.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$ConfigFile = (Join-Path $PSScriptRoot 'config\vcenters.json')
)

$ErrorActionPreference = 'Stop'

# --- Validate config file exists ---
if (-not (Test-Path -Path $ConfigFile)) {
    Write-Error "Configuration file not found: $ConfigFile"
    return
}

$config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
$vaultName = $config.VaultName

# --- Register the vault if it does not already exist ---
$existingVault = Get-SecretVault -Name $vaultName -ErrorAction SilentlyContinue
if (-not $existingVault) {
    Write-Host "Registering SecretStore vault '$vaultName'..." -ForegroundColor Cyan

    # Configure SecretStore for unattended use (no password prompt on access)
    # The vault is still encrypted via Windows DPAPI for the current user
    $storeConfig = @{
        Authentication  = 'None'
        PasswordTimeout = -1
        Interaction     = 'None'
        Confirm         = $false
    }
    Set-SecretStoreConfiguration @storeConfig

    Register-SecretVault -Name $vaultName -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault
    Write-Host "Vault '$vaultName' registered successfully." -ForegroundColor Green
}
else {
    Write-Host "Vault '$vaultName' already exists. Updating credentials..." -ForegroundColor Yellow
}

# --- Prompt for and store credentials for each vCenter ---
foreach ($vc in $config.VCenters) {
    Write-Host "`nEnter credentials for vCenter: $($vc.Name)" -ForegroundColor Cyan
    Write-Host "  Secret will be stored as: $($vc.SecretName)" -ForegroundColor Gray

    $credential = Get-Credential -Message "Credentials for $($vc.Name)"

    if ($null -eq $credential) {
        Write-Warning "Skipped $($vc.Name) - no credential provided."
        continue
    }

    Set-Secret -Name $vc.SecretName -Secret $credential -Vault $vaultName
    Write-Host "  Credential stored for $($vc.Name)." -ForegroundColor Green
}

Write-Host "`nSetup complete. Stored credentials for $($config.VCenters.Count) vCenter(s) in vault '$vaultName'." -ForegroundColor Green
Write-Host "You can now run Get-AllHostInventory.ps1 and Get-AllVMInventory.ps1 unattended." -ForegroundColor Gray
